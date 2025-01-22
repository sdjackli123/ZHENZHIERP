VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw1135 
   BackColor       =   &H00C0E0FF&
   Caption         =   "凭证制作"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form35"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "检平"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "Text6"
      Top             =   4200
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   375
      Left            =   6240
      Top             =   10560
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Caption         =   "Adodc14"
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   5280
      Top             =   10560
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
      Caption         =   "Adodc13"
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   375
      Left            =   5280
      Top             =   10560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   375
      Left            =   7680
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Adodc11"
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   7560
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
      Height          =   330
      Left            =   7560
      Top             =   10560
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
      Height          =   330
      Left            =   4920
      Top             =   10560
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
      Left            =   4920
      Top             =   10560
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
      Left            =   4800
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
      Left            =   4560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   5040
      Top             =   10440
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Left            =   4560
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5640
      Top             =   10560
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
      Left            =   6120
      Top             =   10440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Bindings        =   "Formw1135.frx":0000
      Height          =   2895
      Left            =   120
      TabIndex        =   78
      Top             =   6000
      Width           =   14895
      _cx             =   26273
      _cy             =   5106
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
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formw1135.frx":0015
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   73
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Index           =   0
      Left            =   12000
      TabIndex        =   52
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   0
      Left            =   8760
      TabIndex        =   51
      Top             =   2280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Index           =   0
      Left            =   6000
      TabIndex        =   50
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   0
      Left            =   3720
      TabIndex        =   49
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw1135.frx":002A
      Height          =   330
      Index           =   0
      Left            =   1440
      TabIndex        =   48
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "摘要"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ItemData        =   "Formw1135.frx":003F
      Left            =   9120
      List            =   "Formw1135.frx":0049
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12720
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2160
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
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
      ItemData        =   "Formw1135.frx":005D
      Left            =   4920
      List            =   "Formw1135.frx":006D
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Text            =   "Text5"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Text            =   "Text5"
      Top             =   3720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   422969345
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   44
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   422969345
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3000
      TabIndex        =   45
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   422969345
      CurrentDate     =   39883
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw1135.frx":0099
      Height          =   330
      Index           =   1
      Left            =   1440
      TabIndex        =   53
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "摘要"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw1135.frx":00AE
      Height          =   330
      Index           =   2
      Left            =   1440
      TabIndex        =   54
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "摘要"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw1135.frx":00C3
      Height          =   330
      Index           =   3
      Left            =   1440
      TabIndex        =   55
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "摘要"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formw1135.frx":00D8
      Height          =   330
      Index           =   4
      Left            =   1440
      TabIndex        =   56
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "摘要"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   1
      Left            =   3720
      TabIndex        =   57
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   2
      Left            =   3720
      TabIndex        =   58
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   3
      Left            =   3720
      TabIndex        =   59
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Index           =   4
      Left            =   3720
      TabIndex        =   60
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Index           =   1
      Left            =   6000
      TabIndex        =   61
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Index           =   2
      Left            =   6000
      TabIndex        =   62
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Index           =   3
      Left            =   6000
      TabIndex        =   63
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Index           =   4
      Left            =   6000
      TabIndex        =   64
      Top             =   3720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   8760
      TabIndex        =   65
      Top             =   2640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   8760
      TabIndex        =   66
      Top             =   3000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   8760
      TabIndex        =   67
      Top             =   3360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   4
      Left            =   8760
      TabIndex        =   68
      Top             =   3720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Index           =   1
      Left            =   12000
      TabIndex        =   69
      Top             =   2640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Index           =   2
      Left            =   12000
      TabIndex        =   70
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Index           =   3
      Left            =   12000
      TabIndex        =   71
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Index           =   4
      Left            =   12000
      TabIndex        =   72
      Top             =   3720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formw1135.frx":00ED
      Height          =   330
      Index           =   1
      Left            =   4680
      TabIndex        =   74
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formw1135.frx":0102
      Height          =   330
      Index           =   2
      Left            =   7320
      TabIndex        =   75
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formw1135.frx":0117
      Height          =   330
      Index           =   3
      Left            =   9840
      TabIndex        =   76
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo7"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Bindings        =   "Formw1135.frx":012C
      Height          =   330
      Index           =   4
      Left            =   12840
      TabIndex        =   77
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "MC"
      Text            =   "DataCombo7"
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
      Left            =   1320
      TabIndex        =   47
      Top             =   240
      Width           =   1575
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
      Left            =   3000
      TabIndex        =   46
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   43
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   42
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   41
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   40
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   39
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " 记 账 凭 证"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   5880
      TabIndex        =   38
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   14880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   14880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   14880
      X2              =   14880
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   240
      X2              =   14880
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "具体分类："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   37
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   14880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   14880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line8 
      X1              =   240
      X2              =   14880
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
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
      Left            =   240
      TabIndex        =   36
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "凭证号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11640
      TabIndex        =   34
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "固定号            方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   6
      Left            =   14880
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.Line Line9 
      Index           =   0
      X1              =   3840
      X2              =   8040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "总 账 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   30
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " 明 细 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   29
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Line Line10 
      X1              =   5880
      X2              =   5880
      Y1              =   1800
      Y2              =   4080
   End
   Begin VB.Line Line11 
      X1              =   3720
      X2              =   3720
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line12 
      X1              =   8520
      X2              =   8520
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   240
      X2              =   14880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line13 
      X1              =   5640
      X2              =   9840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line14 
      X1              =   240
      X2              =   14880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "附原始单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   28
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "会计主管"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   27
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "记账"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   26
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "复核"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   6840
      TabIndex        =   25
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出纳"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   9360
      TabIndex        =   24
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "制单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   12360
      TabIndex        =   23
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "凭证类型："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   13
      Left            =   3840
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "贷方金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   11640
      TabIndex        =   21
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "借方金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Line Line9 
      Index           =   1
      X1              =   8280
      X2              =   14880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line16 
      X1              =   11760
      X2              =   11760
      Y1              =   1800
      Y2              =   4080
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      TabIndex        =   19
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "原始单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   15
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "合     计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "摘           要"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   33
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "      发         生         金          额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   5
      Left            =   8760
      TabIndex        =   32
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "      会     计     科     目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   3
      Left            =   3840
      TabIndex        =   79
      Top             =   1440
      Width           =   4215
   End
End
Attribute VB_Name = "Formw1135"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rq1 As ADODB.Parameter
Dim rq2 As ADODB.Parameter
Dim bh As ADODB.Parameter
Public PZH As String: Dim SZSZ(4) As Integer

Private Sub Combo1_Click()
'On Error Resume Next
If Combo2.Text = "付款凭证" Then

If Combo1.Text = "现金" Then

Label7.Caption = "2"
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   
    

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "fkpzhx"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    

If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "2-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "2-1"
End If
Text2.Text = PZH
End If

If Combo1.Text = "银行存款" Then

Label7.Caption = "4"
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   
    

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "fkpzhy"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    

If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "4-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "4-1"
End If
Text2.Text = PZH
End If

End If

If Combo2.Text = "收款凭证" Then


If Combo1.Text = "现金" Then
Label7.Caption = "1"
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   
    

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "skpzhx"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    

If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "1-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "1-1"
End If
Text2.Text = PZH
End If

If Combo1.Text = "银行存款" Then


Label7.Caption = "3"
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   
    

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "fkpzhy"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    

If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "3-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "3-1"
End If
Text2.Text = PZH
End If

End If
End Sub



Private Sub Combo2_Click()
On Error Resume Next
If Combo2.Text = "转账凭证" Then
Text2.Text = ""
For i = 0 To 4
DataCombo2(i).Enabled = True
DataCombo2(i).Text = ""
DataCombo3(i).Enabled = False
DataCombo3(i).Text = ""
Next
Label7.Caption = "5"


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "zzpzh"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    
If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "5-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "5-1"
End If
Text2.Text = PZH
End If

If Combo2.Text = "付款凭证" Then
Text2.Text = ""
For i = 0 To 4
DataCombo2(i).Text = ""
DataCombo2(i).Enabled = True
DataCombo3(i).Text = ""
DataCombo3(i).Enabled = True
Next
End If

If Combo2.Text = "收款凭证" Then
Text2.Text = ""
For i = 0 To 4
DataCombo2(i).Text = ""
DataCombo2(i).Enabled = False
DataCombo3(i).Text = ""
DataCombo3(i).Enabled = False

Next
End If

If Combo2.Text = "成本凭证" Then
Text2.Text = ""
For i = 0 To 4
DataCombo2(i).Enabled = True
DataCombo2(i).Text = ""
DataCombo3(i).Enabled = False
DataCombo3(i).Text = ""
Next
Label7.Caption = "S"

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库

Set param = g_Cmd.CreateParameter("rq1", adDate, adParamInput, 8, DTPicker1.value)
g_Cmd.Parameters.Append param

Set param = g_Cmd.CreateParameter("rq2", adDate, adParamInput, 8, DTPicker2.value)
g_Cmd.Parameters.Append param
    
Set param = g_Cmd.CreateParameter("bh", adInteger, adParamOutput)
g_Cmd.Parameters.Append param
    
   
    

    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "cbpzh"      ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
    

If Val(g_Cmd.Parameters("bh").value) > 0 Then
PZH = "S-" + Trim(Val(g_Cmd.Parameters("bh").value) + 1)
Else
PZH = "S-1"
End If
Text2.Text = PZH
End If

Combo1.Text = ""
End Sub

Private Sub Command1_Click()
For i = 0 To 4
DataCombo1(i).Text = ""
DataCombo2(i).Text = ""
DataCombo3(i).Text = ""
DataCombo4(i).Text = ""
DataCombo5(i).Text = ""
DataCombo7(i).Text = ""
Next
Text3.Text = ""
Combo1.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()

If Val(Text4.Text) = 0 Or Val(Text6.Text) = 0 Then
MsgBox ("借贷金额有误")
Exit Sub
End If

If DataCombo7(4).Text = "" Then
MsgBox ("请输入制单员")
Exit Sub
End If


If Val(Text4.Text) <> Val(Text6.Text) Then
MsgBox ("借贷不平，不能保存！")
Exit Sub
End If

If Combo2.Text = "转账凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Adodc2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
For i = 0 To 4
If DataCombo2(i).Text <> "" Then
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DataCombo1(i).Text
Adodc2.Recordset.Fields(1) = DataCombo2(i).Text
Adodc2.Recordset.Fields(2) = DataCombo3(i).Text
Adodc2.Recordset.Fields(3) = DataCombo4(i).Text
Adodc2.Recordset.Fields(4) = DataCombo5(i).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(i).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
End If
Next
Adodc2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If
'''''''''''''''''''''''''''''''''''''''''''''

If Combo2.Text = "付款凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Adodc2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLFKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
For i = 0 To 4
If DataCombo2(i).Text <> "" Then
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DataCombo1(i).Text
Adodc2.Recordset.Fields(1) = DataCombo2(i).Text
Adodc2.Recordset.Fields(2) = DataCombo3(i).Text
Adodc2.Recordset.Fields(3) = DataCombo4(i).Text
Adodc2.Recordset.Fields(4) = DataCombo5(i).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(i).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
End If
Next
Adodc2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLFKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

If Combo2.Text = "成本凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Adodc2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSCCB.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
For i = 0 To 4
If DataCombo2(i).Text <> "" Then
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DataCombo1(i).Text
Adodc2.Recordset.Fields(1) = DataCombo2(i).Text
Adodc2.Recordset.Fields(2) = DataCombo3(i).Text
Adodc2.Recordset.Fields(3) = DataCombo4(i).Text
Adodc2.Recordset.Fields(4) = DataCombo5(i).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(i).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
End If
Next
Adodc2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSCCB.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

If Combo2.Text = "收款凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Adodc2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
For i = 0 To 4
If DataCombo2(i).Text <> "" Then
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DataCombo1(i).Text
Adodc2.Recordset.Fields(1) = DataCombo2(i).Text
Adodc2.Recordset.Fields(2) = DataCombo3(i).Text
Adodc2.Recordset.Fields(3) = DataCombo4(i).Text
Adodc2.Recordset.Fields(4) = DataCombo5(i).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(i).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
End If
Next
Adodc2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

End Sub

Private Sub Command4_Click()


If DataCombo7(4).Text = "" Then
MsgBox ("请输入制单员")
Exit Sub
End If


If Combo2.Text = "转账凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If MsgBox("确认修改吗？", vbYesNo) = vbNo Then Exit Sub
Adodc2.Recordset.Fields(0) = DataCombo1(0).Text
Adodc2.Recordset.Fields(1) = DataCombo2(0).Text
Adodc2.Recordset.Fields(2) = DataCombo3(0).Text
Adodc2.Recordset.Fields(3) = DataCombo4(0).Text
Adodc2.Recordset.Fields(4) = DataCombo5(0).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(0).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
Adodc2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
MsgBox ("修改成功！")
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLZZPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If
'''''''''''''''''''''''''''''''''''''''''''''

If Combo2.Text = "付款凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If MsgBox("确认修改吗？", vbYesNo) = vbNo Then Exit Sub
Adodc2.Recordset.Fields(0) = DataCombo1(0).Text
Adodc2.Recordset.Fields(1) = DataCombo2(0).Text
Adodc2.Recordset.Fields(2) = DataCombo3(0).Text
Adodc2.Recordset.Fields(3) = DataCombo4(0).Text
Adodc2.Recordset.Fields(4) = DataCombo5(0).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(0).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
Adodc2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLFKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
MsgBox ("修改成功！")
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLFKPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If

If Combo2.Text = "收款凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If MsgBox("确认修改吗？", vbYesNo) = vbNo Then Exit Sub
Adodc2.Recordset.Fields(0) = DataCombo1(0).Text
Adodc2.Recordset.Fields(1) = DataCombo2(0).Text
Adodc2.Recordset.Fields(2) = DataCombo3(0).Text
Adodc2.Recordset.Fields(3) = DataCombo4(0).Text
Adodc2.Recordset.Fields(4) = DataCombo5(0).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(0).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
Adodc2.RecordSource = "SELECT * FROM CLSKPZ WHERE  日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
MsgBox ("修改成功！")
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLSKPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If


If Combo2.Text = "成本凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DataCombo7(4).Text = "" Or DataCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If MsgBox("确认修改吗？", vbYesNo) = vbNo Then Exit Sub
Adodc2.Recordset.Fields(0) = DataCombo1(0).Text
Adodc2.Recordset.Fields(1) = DataCombo2(0).Text
Adodc2.Recordset.Fields(2) = DataCombo3(0).Text
Adodc2.Recordset.Fields(3) = DataCombo4(0).Text
Adodc2.Recordset.Fields(4) = DataCombo5(0).Text
Adodc2.Recordset.Fields(5) = Text2.Text
Adodc2.Recordset.Fields(6) = DTPicker3.value
Adodc2.Recordset.Fields(7) = Text5(0).Text
Adodc2.Recordset.Fields(8) = DataCombo7(1).Text
Adodc2.Recordset.Fields(9) = DataCombo7(2).Text
Adodc2.Recordset.Fields(10) = DataCombo7(4).Text
Adodc2.Recordset.Fields(11) = Text3.Text
Adodc2.Recordset.Update
Adodc2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSCCB.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
MsgBox ("修改成功！")
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLSCCB where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If

End Sub

Private Sub Command5_Click()
On Error Resume Next
If MsgBox("删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
If Combo2.Text = "转账凭证" Then
Adodc2.Recordset.Delete
Adodc2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from clzzpz where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If
'''''''''''''''''''''''''''''''''''''''''''''
If Combo2.Text = "付款凭证" Then
Adodc2.Recordset.Delete
Adodc2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLFKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLFKPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If

If Combo2.Text = "收款凭证" Then
Adodc2.Recordset.Delete
Adodc2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLSKPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If

If Combo2.Text = "成本凭证" Then
Adodc2.Recordset.Delete
Adodc2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLSCCB where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平")
End If
End If

End Sub

Private Sub Command6_Click()
Adodc10.RecordSource = "select sum(cast(借方金额 as real)),sum(cast(贷方金额 as real)) from CLZZPZ where 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc10.Refresh
If Adodc10.Recordset.Fields(0) <> Adodc10.Recordset.Fields(1) Then
MsgBox ("借贷不平" + Trim(Adodc10.Recordset.Fields(0) - Adodc10.Recordset.Fields(1)))
End If
End Sub

Private Sub DataCombo4_Change(Index As Integer)
Select Case Index
       Case Index
Text4.Text = Format(Val(DataCombo4(0).Text) + Val(DataCombo4(1).Text) + Val(DataCombo4(2).Text) + Val(DataCombo4(3).Text) + Val(DataCombo4(4).Text), "#0.00")
End Select
End Sub

Private Sub DataCombo5_Change(Index As Integer)
Select Case Index
       Case Index
Text6.Text = Format(Val(DataCombo5(0).Text) + Val(DataCombo5(1).Text) + Val(DataCombo5(2).Text) + Val(DataCombo5(3).Text) + Val(DataCombo5(4).Text), "#0.00")
End Select
End Sub

Private Sub DTPicker3_Change()
Text1.Text = Month(DTPicker3.value)
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = Month(DTPicker3.value)
End Sub

Private Sub Form_Load()

On Error Resume Next
Combo1.Text = ""
DTPicker3.value = Date
Text7.Text = Date
Text1.Text = Month(DTPicker3.value)
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.RecordSource = "select * from rqsd where 月份='" & Text1.Text & "'"
Adodc14.Refresh
If Adodc14.Recordset.EOF Then
Exit Sub
Else
DTPicker1.value = Adodc14.Recordset.Fields(0)
DTPicker2.value = Adodc14.Recordset.Fields(1)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from GYS  GROUP BY 简称"
Adodc1.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select CWZY.摘要 from CWZY  GROUP BY CWZY.摘要"
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 科目名称 from CWMC WHERE 科目方向='借' GROUP BY 科目名称"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select 科目名称 from CWMC WHERE 科目方向='贷' GROUP BY 科目名称"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 简称 from khzl  GROUP BY 简称"
Adodc8.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select FHY.MC from FHY GROUP BY FHY.MC"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Label7.Caption = ""
For i = 0 To 4
DataCombo1(i).Text = ""
DataCombo2(i).Text = ""
DataCombo3(i).Text = ""
DataCombo4(i).Text = ""
DataCombo5(i).Text = ""
'DataCombo6(i).Text = ""
DataCombo7(i).Text = ""
Text5(i).Text = ""
Next
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = "0"
Text6.Text = "0"

'vSFlexGrid1.ColWidth(13) = 0

End Sub
Private Sub DTPicker1_Change()
Text6.Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text6.Text = DTPicker1.value
End Sub


Private Sub Label10_Click(Index As Integer)
Select Case Index
       Case Index
       KMBL = Index
       KMMC = 4
Formw6.Show
End Select

End Sub


Private Sub Label8_Click(Index As Integer)
Select Case Index
       Case Index
       KMBL = Index
       KMMC = 2
Formw6.Show
End Select
End Sub


Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
rs = VSFlexGrid2.Row
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
 DataCombo1(0).Text = Adodc2.Recordset.Fields(0)
 DataCombo2(0).Text = Adodc2.Recordset.Fields(1)
 DataCombo3(0).Text = Adodc2.Recordset.Fields(2)
 DataCombo4(0).Text = Adodc2.Recordset.Fields(3)
 DataCombo5(0).Text = Adodc2.Recordset.Fields(4)
 DTPicker3.value = Adodc2.Recordset.Fields(6)
 Text5(0).Text = Adodc2.Recordset.Fields(7)
 DataCombo7(1).Text = Adodc2.Recordset.Fields(8)
 DataCombo7(2).Text = Adodc2.Recordset.Fields(9)
 DataCombo7(4).Text = Adodc2.Recordset.Fields(10)
End Sub

Private Sub Text1_Change()
On Error Resume Next
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.RecordSource = "select * from rqsd where 月份='" & Text1.Text & "'"
Adodc14.Refresh
If Adodc14.Recordset.EOF Then
Exit Sub
Else
DTPicker1.value = Adodc14.Recordset.Fields(0)
DTPicker2.value = Adodc14.Recordset.Fields(1)
End If
End Sub

Private Sub Text2_Change()
If Combo2.Text = "转账凭证" Then
Adodc2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLZZPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If
'''''''''''''''''''''''''''''''''''''''''''''
If Combo2.Text = "付款凭证" Then
Adodc2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLFKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

If Combo2.Text = "收款凭证" Then
Adodc2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSKPZ.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

If Combo2.Text = "成本凭证" Then
Adodc2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND CLSCCB.凭证号='" & Text2.Text & "'"
Adodc2.Refresh
End If

End Sub

