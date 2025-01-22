VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formw332 
   BackColor       =   &H00C0E0FF&
   Caption         =   "凭证操作"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form32"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
      Height          =   1095
      Left            =   600
      TabIndex        =   47
      Top             =   8160
      Width           =   13815
      _cx             =   24368
      _cy             =   1931
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
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Height          =   1215
      Left            =   5520
      TabIndex        =   46
      Top             =   1680
      Width           =   3495
      _cx             =   6165
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
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Height          =   1695
      Left            =   360
      TabIndex        =   45
      Top             =   5280
      Width           =   13335
      _cx             =   23521
      _cy             =   2990
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
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   7560
      Top             =   9960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
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
      Left            =   7320
      Top             =   10200
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      Left            =   7920
      Top             =   9480
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
      Left            =   7800
      Top             =   9600
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
      Left            =   7920
      Top             =   9840
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
      Left            =   8160
      Top             =   9360
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
      Left            =   7920
      Top             =   9600
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
      Left            =   7800
      Top             =   9720
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
      Height          =   330
      Left            =   8040
      Top             =   10080
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
      Left            =   8280
      Top             =   9480
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
      Left            =   8520
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   10200
      TabIndex        =   44
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   1560
      TabIndex        =   43
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   1560
      TabIndex        =   42
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.ListBox List1 
      Height          =   1110
      ItemData        =   "Formw332.frx":0000
      Left            =   9360
      List            =   "Formw332.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   41
      Top             =   1680
      Width           =   5535
   End
   Begin VB.ListBox List2 
      Height          =   1110
      ItemData        =   "Formw332.frx":0004
      Left            =   9360
      List            =   "Formw332.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   40
      Top             =   1680
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "批复"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "批否"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3120
      Width           =   975
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证复核"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formw332.frx":0008
      Height          =   1095
      Left            =   240
      TabIndex        =   33
      Top             =   7080
      Visible         =   0   'False
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   12720
      TabIndex        =   29
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "成批"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "单个"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "登账"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "凭证打印"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "复核确定"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1800
      Top             =   0
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   10320
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "Formw332.frx":001C
      Left            =   1560
      List            =   "Formw332.frx":002C
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Formw332.frx":0058
      Left            =   3480
      List            =   "Formw332.frx":0062
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw332.frx":006E
      Height          =   330
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo2"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   79691777
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      Format          =   79691777
      CurrentDate     =   36892
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw332.frx":0082
      Height          =   1455
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   9
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      GridLines       =   2
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw332.frx":0096
      Height          =   330
      Left            =   2880
      TabIndex        =   15
      Top             =   2160
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "凭证号"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formw332.frx":00AA
      Height          =   855
      Left            =   5400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   5
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "记录号 "
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   79691777
      CurrentDate     =   39883
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formw332.frx":00BE
      Height          =   360
      Left            =   11640
      TabIndex        =   32
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   635
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   35
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "审核无误可进行登账处理"
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
      Left            =   9360
      TabIndex        =   28
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   26
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   25
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "选择日期范围"
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
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选择凭证"
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
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "请选复核员"
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
      Left            =   360
      TabIndex        =   22
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "凭证类别："
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
      Left            =   360
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
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
      Left            =   3480
      TabIndex        =   20
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
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
      Left            =   3480
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "审核通过"
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
      Left            =   3480
      TabIndex        =   18
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "Formw332"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DD, BAR, c, r As Integer: Public K1, K2 As String

Private Sub Combo1_Click()
'On Error Resume Next
If Combo1.Text = "转账凭证" Then
Data2.RecordSource = "select CLZZPZ.凭证号 from CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by CLZZPZ.凭证号"
Data2.Refresh
Data6.RecordSource = "SELECT CLZZPZ.日期,CLZZPZ.凭证号,CLZZPZ.审核确认,CLZZPZ.记账标记 FROM CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLZZPZ.日期,CLZZPZ.凭证号,CLZZPZ.审核确认,CLZZPZ.记账标记 ORDER BY CLZZPZ.日期,VAL(MID(CLZZPZ.凭证号,3))"
Data6.Refresh
End If

If Combo1.Text = "付款凭证" Then
Data2.RecordSource = "select CLFKPZ.凭证号 from CLFKPZ WHERE CLFKPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by CLFKPZ.凭证号"
Data2.Refresh
Data6.RecordSource = "SELECT CLFKPZ.日期,CLFKPZ.凭证号,CLFKPZ.审核确认,CLFKPZ.记账标记 FROM CLFKPZ WHERE CLFKPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLFKPZ.日期,CLFKPZ.凭证号,CLFKPZ.审核确认,CLFKPZ.记账标记 ORDER BY CLFKPZ.日期,VAL(MID(CLFKPZ.凭证号,3))"
Data6.Refresh
End If

If Combo1.Text = "收款凭证" Then
Data2.RecordSource = "select CLSKPZ.凭证号 from CLSKPZ WHERE CLSKPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by CLSKPZ.凭证号"
Data2.Refresh
Data6.RecordSource = "SELECT CLSKPZ.日期,CLSKPZ.凭证号,CLSKPZ.审核确认,CLSKPZ.记账标记 FROM CLSKPZ WHERE CLSKPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLSKPZ.日期,CLSKPZ.凭证号,CLSKPZ.审核确认,CLSKPZ.记账标记 ORDER BY CLSKPZ.日期,VAL(MID(CLSKPZ.凭证号,3))"
Data6.Refresh
End If

If Combo1.Text = "成本凭证" Then
Data2.RecordSource = "select CLSCCB.凭证号 from CLSCCB WHERE CLSCCB.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by CLSCCB.凭证号"
Data2.Refresh
Data6.RecordSource = "SELECT CLSCCB.日期,CLSCCB.凭证号,CLSCCB.审核确认,CLSCCB.记账标记 FROM CLSCCB WHERE CLSCCB.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLSCCB.日期,CLSCCB.凭证号,CLSCCB.审核确认,CLSCCB.记账标记 ORDER BY CLSCCB.日期,VAL(MID(CLSCCB.凭证号,3))"
Data6.Refresh
End If

End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Or DBCombo1.Text = "" Then
MsgBox ("请输入凭证类别和凭证号")
Exit Sub
End If
If Data1.Recordset.EOF Then Exit Sub
Call PZDY(Combo1.Text, DBCombo1.Text)
End Sub

Private Sub Command10_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Command2_Click()
If DBCombo1.Text = "" Then
MsgBox ("请输入凭证号")
Exit Sub
End If
If DBCombo2.Text = "" Then
MsgBox ("请输入复核员")
Exit Sub
End If
If Combo1.Text = "转账凭证" Then
Data3.Database.Execute "UPDATE CLZZPZ SET 复核='" & DBCombo2.Text & "',审核确认='" & Combo2.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "' "
Data1.RecordSource = "select CLZZPZ.原始单据,CLZZPZ.摘要,CLZZPZ.借方总账科目,CLZZPZ.贷方总账科目,CLZZPZ.贷方明细科目,CLZZPZ.金额,CLZZPZ.凭证号,CLZZPZ.记账,CLZZPZ.复核,CLZZPZ.制单,CLZZPZ.审核确认 from CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "付款凭证" Then
Data3.Database.Execute "UPDATE CLFKPZ SET 复核='" & DBCombo2.Text & "',审核确认='" & Combo2.Text & "' WHERE 凭证号='" & DBCombo1.Text & "' AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') "
Data1.RecordSource = "select CLFKPZ.原始单据,CLFKPZ.摘要,CLFKPZ.借方总账科目,CLFKPZ.贷方总账科目,CLFKPZ.贷方明细科目,CLFKPZ.金额,CLFKPZ.凭证号,CLFKPZ.记账,CLFKPZ.复核,CLFKPZ.制单,CLFKPZ.审核确认 from CLFKPZ WHERE  日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "收款凭证" Then
Data3.Database.Execute "UPDATE CLSKPZ SET 复核='" & DBCombo2.Text & "',审核确认='" & Combo2.Text & "' WHERE 凭证号='" & DBCombo1.Text & "' AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data1.RecordSource = "select CLSKPZ.原始单据,CLSKPZ.摘要,CLSKPZ.借方总账科目,CLSKPZ.贷方总账科目,CLSKPZ.贷方明细科目,CLSKPZ.金额,CLSKPZ.凭证号,CLSKPZ.记账,CLSKPZ.复核,CLSKPZ.制单,CLSKPZ.审核确认 from CLSKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "成本凭证" Then
Data3.Database.Execute "UPDATE CLSCCB SET 复核='" & DBCombo2.Text & "',审核确认='" & Combo2.Text & "' WHERE 凭证号='" & DBCombo1.Text & "' and 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') "
Data1.RecordSource = "select CLSCCB.原始单据,CLSCCB.摘要,CLSCCB.借方总账科目,CLSCCB.贷方总账科目,CLSCCB.贷方明细科目,CLSCCB.金额,CLSCCB.凭证号,CLSCCB.记账,CLSCCB.复核,CLSCCB.制单,CLSCCB.审核确认 from CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

Data6.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
If Combo1.Text = "转账凭证" Then
Data4.RecordSource = "select * from CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "'  "
Data4.Refresh
End If

If Combo1.Text = "付款凭证" Then
Data4.RecordSource = "select * from CLFKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "' "
Data4.Refresh
End If

If Combo1.Text = "收款凭证" Then
Data4.RecordSource = "select * from CLSKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "' "
Data4.Refresh
End If

If Combo1.Text = "成本凭证" Then
Data4.RecordSource = "select * from CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "' "
Data4.Refresh
End If

End Sub


Private Sub Command5_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command6_Click()
If Option1.Value = True Then
If DBCombo1.Text = "" Then
MsgBox ("请输入凭证号")
Exit Sub
End If

If DBCombo3.Text = "" Then
MsgBox ("请输入登账人")
Exit Sub
End If

Data6.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "' AND 审核确认='已'"
If Data6.Recordset.NoMatch Then
MsgBox ("没有复核，不能登账！")
Exit Sub
End If

Data6.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "' AND 审核确认='已' AND 记账标记='记'"
If Data6.Recordset.NoMatch Then
Else
MsgBox ("已登账，不能重复登账！")
Exit Sub
End If

Data9.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
If Combo1.Text = "转账凭证" Then
Call Chk1
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "收款凭证" Then
Call Chk2
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLSKPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "付款凭证" Then
Call Chk3
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLFKPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "成本凭证" Then
Call Chk4
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLSCCB set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

Else
MsgBox (Combo1.Text + DBCombo1.Text + "已登账")
End If
Call Option1_Click
End If

'''''''''''''''''''''''''''批量登账
If Option2.Value = True Then
If DBCombo3.Text = "" Then
MsgBox ("请输入登账人")
Exit Sub
End If
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
DBCombo1.Text = Trim(Mid(List1.List(i), 1, 10))
Data6.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "' AND 审核确认='已'"
If Data6.Recordset.NoMatch Then
MsgBox ("没有复核，不能登账！")
Exit Sub
End If

Data6.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "' AND 审核确认='已' AND 记账标记='记'"
If Data6.Recordset.NoMatch Then
Else
MsgBox ("已登账，不能重复登账！")
Exit Sub
End If

Data9.Recordset.FindFirst "日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND 凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
If Combo1.Text = "转账凭证" Then
Call Chk1
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "收款凭证" Then
Call Chk2
Data1.Database.Execute "UPDATE CLSKPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "付款凭证" Then
Call Chk3
Data1.Database.Execute "UPDATE CLFKPZ set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

If Combo1.Text = "成本凭证" Then
Call Chk4
Data1.Database.Execute "UPDATE CLSCCB set 记账标记='记',记账='" & DBCombo3.Text & "' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "'"
Data6.Refresh
End If

End If

End If
Next
Call Option2_Click
End If
End Sub

Private Sub Command7_Click()
Data1.Database.Execute "UPDATE CLZZPZ SET 借方总账科目='' WHERE 借方总账科目=NULL"
Data1.Database.Execute "UPDATE CLZZPZ SET 借方明细科目='' WHERE 借方明细科目=NULL"
Data1.Database.Execute "UPDATE CLZZPZ SET 贷方总账科目='' WHERE 贷方总账科目=NULL"
Data1.Database.Execute "UPDATE CLZZPZ SET 贷方明细科目='' WHERE 贷方明细科目=NULL"

Data1.Database.Execute "UPDATE CLFKPZ SET 借方总账科目='' WHERE 借方总账科目=NULL"
Data1.Database.Execute "UPDATE CLFKPZ SET 借方明细科目='' WHERE 借方明细科目=NULL"
Data1.Database.Execute "UPDATE CLFKPZ SET 贷方总账科目='' WHERE 贷方总账科目=NULL"
Data1.Database.Execute "UPDATE CLFKPZ SET 贷方明细科目='' WHERE 贷方明细科目=NULL"

Data1.Database.Execute "UPDATE CLSKPZ SET 借方总账科目='' WHERE 借方总账科目=NULL"
Data1.Database.Execute "UPDATE CLSKPZ SET 借方明细科目='' WHERE 借方明细科目=NULL"
Data1.Database.Execute "UPDATE CLSKPZ SET 贷方总账科目='' WHERE 贷方总账科目=NULL"
Data1.Database.Execute "UPDATE CLSKPZ SET 贷方明细科目='' WHERE 贷方明细科目=NULL"

Data1.Database.Execute "UPDATE CLSCCB SET 借方总账科目='' WHERE 借方总账科目=NULL"
Data1.Database.Execute "UPDATE CLSCCB SET 借方明细科目='' WHERE 借方明细科目=NULL"
Data1.Database.Execute "UPDATE CLSCCB SET 贷方总账科目='' WHERE 贷方总账科目=NULL"
Data1.Database.Execute "UPDATE CLSCCB SET 贷方明细科目='' WHERE 贷方明细科目=NULL"

MsgBox ("凭证复核成功！可进行相关操作！")
Label8.Caption = "凭证复核成功！可进行相关操作！"

End Sub

Private Sub Command8_Click()
Formw1133.Text1.Text = Text3.Text
Formw1133.Combo1.Text = Combo1.Text
Formw1133.Show
End Sub

Private Sub Command9_Click()
Formw1131.Text1.Text = Text3.Text
Formw1131.Combo1.Text = Combo1.Text
Formw1131.Show
End Sub

Private Sub DBCombo1_Change()
If Combo1.Text = "转账凭证" Then
Data1.RecordSource = "select CLZZPZ.原始单据,CLZZPZ.摘要,CLZZPZ.借方总账科目,CLZZPZ.借方明细科目,CLZZPZ.贷方总账科目,CLZZPZ.贷方明细科目,CLZZPZ.金额,CLZZPZ.凭证号,CLZZPZ.记账,CLZZPZ.复核,CLZZPZ.制单,CLZZPZ.审核确认,日期 from CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "付款凭证" Then
Data1.RecordSource = "select CLFKPZ.原始单据,CLFKPZ.摘要,CLFKPZ.借方总账科目,借方明细科目,CLFKPZ.贷方总账科目,CLFKPZ.贷方明细科目,CLFKPZ.金额,CLFKPZ.凭证号,CLFKPZ.记账,CLFKPZ.复核,CLFKPZ.制单,CLFKPZ.审核确认,日期 from CLFKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "收款凭证" Then
Data1.RecordSource = "select CLSKPZ.原始单据,CLSKPZ.摘要,CLSKPZ.借方总账科目,借方明细科目,CLSKPZ.贷方总账科目,CLSKPZ.贷方明细科目,CLSKPZ.金额,CLSKPZ.凭证号,CLSKPZ.记账,CLSKPZ.复核,CLSKPZ.制单,CLSKPZ.审核确认,日期 from CLSKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "成本凭证" Then
Data1.RecordSource = "select CLSCCB.原始单据,CLSCCB.摘要,CLSCCB.借方总账科目,借方明细科目,CLSCCB.贷方总账科目,CLSCCB.贷方明细科目,CLSCCB.金额,CLSCCB.凭证号,CLSCCB.记账,CLSCCB.复核,CLSCCB.制单,CLSCCB.审核确认,日期 from CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

End Sub

Private Sub DBCombo1_Click(Area As Integer)
If Combo1.Text = "转账凭证" Then
Data1.RecordSource = "select CLZZPZ.原始单据,CLZZPZ.摘要,CLZZPZ.借方总账科目,借方明细科目,CLZZPZ.贷方总账科目,CLZZPZ.贷方明细科目,CLZZPZ.金额,CLZZPZ.凭证号,CLZZPZ.记账,CLZZPZ.复核,CLZZPZ.制单,CLZZPZ.审核确认,日期 from CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "付款凭证" Then
Data1.RecordSource = "select CLFKPZ.原始单据,CLFKPZ.摘要,CLFKPZ.借方总账科目,借方明细科目,CLFKPZ.贷方总账科目,CLFKPZ.贷方明细科目,CLFKPZ.金额,CLFKPZ.凭证号,CLFKPZ.记账,CLFKPZ.复核,CLFKPZ.制单,CLFKPZ.审核确认,日期 from CLFKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "收款凭证" Then
Data1.RecordSource = "select CLSKPZ.原始单据,CLSKPZ.摘要,CLSKPZ.借方总账科目,借方明细科目,CLSKPZ.贷方总账科目,CLSKPZ.贷方明细科目,CLSKPZ.金额,CLSKPZ.凭证号,CLSKPZ.记账,CLSKPZ.复核,CLSKPZ.制单,CLSKPZ.审核确认,日期 from CLSKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

If Combo1.Text = "成本凭证" Then
Data1.RecordSource = "select CLSCCB.原始单据,CLSCCB.摘要,CLSCCB.借方总账科目,借方明细科目,CLSCCB.贷方总账科目,CLSCCB.贷方明细科目,CLSCCB.金额,CLSCCB.凭证号,CLSCCB.记账,CLSCCB.复核,CLSCCB.制单,CLSCCB.审核确认,日期 from CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "' "
Data1.Refresh
End If

End Sub

Private Sub DBCOMBO1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DTPicker1_Change()
Data11.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data11.RecordSource = "select * from RQSD where cdate('" & DTPicker1.Value & "') between 起始日期 and 结束日期"
Data11.Refresh
If Data11.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data11.Recordset.Fields(0)
K2 = Data11.Recordset.Fields(1)
Text3.Text = Data11.Recordset.Fields(2)
End If


Text1.Text = K1
Text2.Text = K2
DTPicker3.Value = K1
DTPicker4.Value = K2
End Sub

Private Sub DTPicker1_CloseUp()
Data11.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data11.RecordSource = "select * from RQSD where cdate('" & DTPicker1.Value & "') between 起始日期 and 结束日期"
Data11.Refresh
If Data11.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data11.Recordset.Fields(0)
K2 = Data11.Recordset.Fields(1)
Text3.Text = Data11.Recordset.Fields(2)
End If


Text1.Text = K1
Text2.Text = K2
DTPicker3.Value = K1
DTPicker4.Value = K2
End Sub

Private Sub DTPicker3_Change()
Text1.Text = DTPicker3.Value
End Sub

Private Sub DTPicker3_CloseUp()
Text1.Text = DTPicker3.Value
Text1.SetFocus
End Sub

Private Sub DTPicker4_Change()
Text2.Text = DTPicker4.Value
End Sub

Private Sub DTPicker4_CloseUp()
Text2.Text = DTPicker4.Value
Text2.SetFocus
End Sub

Private Sub Form_Load()
Combo1.Text = ""
Combo2.Text = ""
Option1.Value = True
DTPicker1.Value = Date
DBCombo3.Text = ""
Label8.Caption = "请先进行凭证复核！"

Data11.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data11.RecordSource = "select * from RQSD where cdate('" & DTPicker1.Value & "') between 起始日期 and 结束日期"
Data11.Refresh
If Data11.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data11.Recordset.Fields(0)
K2 = Data11.Recordset.Fields(1)
Text3.Text = Data11.Recordset.Fields(2)
End If


Text1.Text = K1
Text2.Text = K2
DTPicker3.Value = K1
DTPicker4.Value = K2

Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data1.Refresh

Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data2.Refresh

Data3.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data4.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"

Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data5.RecordSource = "select FHY.MC from FHY GROUP BY FHY.MC"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data7.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"

Data8.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.mdb"

Data9.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data9.RecordSource = "SELECT * FROM PZDZ WHERE PZDZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"

ProgressBar1.Visible = False
Timer1.Enabled = False
DBCombo1.Text = ""
DBCombo2.Text = ""

MSFlexGrid1.ColWidth(1) = 1200
MSFlexGrid1.ColWidth(2) = 1200
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1200
MSFlexGrid1.ColWidth(7) = 1200

MSFlexGrid3.ColWidth(0) = 100
MSFlexGrid3.ColWidth(1) = 1000
MSFlexGrid3.ColWidth(2) = 600
MSFlexGrid3.ColWidth(3) = 600
MSFlexGrid3.ColWidth(4) = 600
End Sub

Private Sub MSFlexGrid3_DBLClick()
rs = MSFlexGrid3.Row
If Data6.Recordset.EOF Then Exit Sub
Data6.Recordset.MoveFirst
Data6.Recordset.Move rs - 1
DBCombo1.Text = Data6.Recordset.Fields(1)
End Sub


Private Sub Option1_Click()
'On Error Resume Next
If Option1.Value = True Then
List1.Visible = False
List2.Visible = True
Command5.Visible = False
Command10.Visible = False
If Combo1.Text = "转账凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLZZPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List2.Clear
Do While Not Data4.Recordset.EOF
List2.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "付款凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLFKPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List2.Clear
Do While Not Data4.Recordset.EOF
List2.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "收款凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLSKPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List2.Clear
Do While Not Data4.Recordset.EOF
List2.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "成本凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLSCCB WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List2.Clear
Do While Not Data4.Recordset.EOF
List2.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
End If
End Sub

Private Sub Option2_Click()
On Error Resume Next
If Option2.Value = True Then
Command5.Visible = True
Command10.Visible = True
List2.Visible = False
List1.Visible = True
If Combo1.Text = "转账凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLZZPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List1.Clear
Do While Not Data4.Recordset.EOF
List1.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "付款凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLFKPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List1.Clear
Do While Not Data4.Recordset.EOF
List1.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "收款凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLSKPZ WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List1.Clear
Do While Not Data4.Recordset.EOF
List1.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
If Combo1.Text = "成本凭证" Then
Data4.RecordSource = "select 凭证号,日期 from CLSCCB WHERE 审核确认='已' AND (记账标记=NULL OR 记账标记<>'记') AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') group by 凭证号,日期 ORDER BY VAL(MID(凭证号,3))"
Data4.Refresh
If Data4.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Data4.Recordset.MoveFirst
List1.Clear
Do While Not Data4.Recordset.EOF
List1.AddItem Data4.Recordset.Fields(0) + Space(10) + Trim(Data4.Recordset.Fields(1))
Data4.Recordset.MoveNext
Loop
End If
End If
End Sub

Private Sub Timer1_Timer()
If BAR = 100 Then
DataEnvironment6.CLZZPZ DBCombo1.Text
DataReport16.Show 1
DataEnvironment6.rsCLZZPZ.Close
ProgressBar1.Visible = False
Timer1.Enabled = False
Exit Sub
Else
ProgressBar1.Value = BAR
BAR = BAR + 1
End If
End Sub
Private Sub MSFlex_DBLClick()
With MSFlexGrid1
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex_DBLClick
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid1.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid1.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid1.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
If c = 4 Or c = 7 Then
Data1.Recordset.MoveFirst
Data1.Recordset.Move r - 1
Data1.Recordset.Edit
Data1.Recordset.Fields(c - 1) = Text1111.Text
Data1.Recordset.Update
Text1111.Visible = False
End If
End Sub

Private Sub Chk3()
Data1.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLFKPZ.凭证号='" & DBCombo1.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(2) = Null Or Trim(Data1.Recordset.Fields(2)) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1) + "-" + Data1.Recordset.Fields(2)
End If
Data9.Recordset.Fields(4) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(5) = 0
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update

Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(4) = Null Or Data1.Recordset.Fields(4) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3) + "-" + Data1.Recordset.Fields(4)
End If
Data9.Recordset.Fields(4) = 0
Data9.Recordset.Fields(5) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Chk1()
Data1.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & DBCombo1.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(2) = Null Or Trim(Data1.Recordset.Fields(2)) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1) + "-" + Data1.Recordset.Fields(2)
End If
Data9.Recordset.Fields(4) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(5) = 0
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update

Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(4) = Null Or Data1.Recordset.Fields(4) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3) + "-" + Data1.Recordset.Fields(4)
End If
Data9.Recordset.Fields(4) = 0
Data9.Recordset.Fields(5) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update
Data1.Recordset.MoveNext
Loop
End Sub

Private Sub Chk2()
Data1.RecordSource = "SELECT * FROM CLSKPZ WHERE  日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSKPZ.凭证号='" & DBCombo1.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(2) = Null Or Trim(Data1.Recordset.Fields(2)) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1) + "-" + Data1.Recordset.Fields(2)
End If
Data9.Recordset.Fields(4) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(5) = 0
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update

Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(4) = Null Or Data1.Recordset.Fields(4) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3) + "-" + Data1.Recordset.Fields(4)
End If
Data9.Recordset.Fields(4) = 0
Data9.Recordset.Fields(5) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update
Data1.Recordset.MoveNext
Loop
End Sub
Private Sub Chk4()
Data1.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLSCCB.凭证号='" & DBCombo1.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(2) = Null Or Trim(Data1.Recordset.Fields(2)) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(1) + "-" + Data1.Recordset.Fields(2)
End If
Data9.Recordset.Fields(4) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(5) = 0
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update

Data9.Recordset.AddNew
Data9.Recordset.Fields(0) = Data1.Recordset.Fields(7)
Data9.Recordset.Fields(1) = Data1.Recordset.Fields(6)
Data9.Recordset.Fields(2) = Data1.Recordset.Fields(0)
If Data1.Recordset.Fields(4) = Null Or Data1.Recordset.Fields(4) = "" Then
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3)
Else
Data9.Recordset.Fields(3) = Data1.Recordset.Fields(3) + "-" + Data1.Recordset.Fields(4)
End If
Data9.Recordset.Fields(4) = 0
Data9.Recordset.Fields(5) = Data1.Recordset.Fields(5)
Data9.Recordset.Fields(6) = Data1.Recordset.Fields(11)
Data9.Recordset.Fields(7) = Data1.Recordset.Fields(10)
Data9.Recordset.Fields(8) = Data1.Recordset.Fields(9)
Data9.Recordset.Fields(9) = Combo1.Text
Data9.Recordset.Update
Data1.Recordset.MoveNext
Loop
End Sub


Private Sub DZ1()
Data10.RecordSource = "SELECT CLZZPZ.日期,CLZZPZ.凭证号,CLZZPZ.审核确认,CLZZPZ.记账标记 FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.审核确认='已' AND (CLZZPZ.记账标记='' OR CLZZPZ.记账标记=NULL) GROUP BY CLZZPZ.日期,CLZZPZ.凭证号,CLZZPZ.审核确认,CLZZPZ.记账标记 ORDER BY CLZZPZ.日期,VAL(MID(CLZZPZ.凭证号,3))"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
DBCombo1.Text = Data10.Recordset.Fields(1)
Do While Not Data10.Recordset.EOF
Data9.Recordset.FindFirst "凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
Call Chk1
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & Data10.Recordset.Fields(1) & "'"
Data6.Refresh
Else
MsgBox (DBCombo1.Text + "已登账")
End If
Data10.Recordset.MoveNext
DBCombo1.Text = Data10.Recordset.Fields(1)
Loop
End Sub

Private Sub DZ2()
Data10.RecordSource = "SELECT CLSKPZ.日期,CLSKPZ.凭证号,CLSKPZ.审核确认,CLSKPZ.记账标记 FROM CLSKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLSKPZ.日期,CLSKPZ.凭证号,CLSKPZ.审核确认,CLSKPZ.记账标记 ORDER BY CLSKPZ.日期,VAL(MID(CLSKPZ.凭证号,3))"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
DBCombo1.Text = Data10.Recordset.Fields(1)
Do While Not Data10.Recordset.EOF
Data9.Recordset.FindFirst "凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
Call Chk2
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & Data10.Recordset.Fields(1) & "'"
Data6.Refresh
Else
MsgBox (DBCombo1.Text + "已登账")
End If
Data10.Recordset.MoveNext
DBCombo1.Text = Data10.Recordset.Fields(1)
Loop
End Sub

Private Sub DZ3()
Data10.RecordSource = "SELECT CLFKPZ.日期,CLFKPZ.凭证号,CLFKPZ.审核确认,CLFKPZ.记账标记 FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') GROUP BY CLFKPZ.日期,CLFKPZ.凭证号,CLFKPZ.审核确认,CLFKPZ.记账标记 ORDER BY CLFKPZ.日期,VAL(MID(CLFKPZ.凭证号,3))"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
DBCombo1.Text = Data10.Recordset.Fields(1)
Do While Not Data10.Recordset.EOF
Data9.Recordset.FindFirst "凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
Call Chk3
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & Data10.Recordset.Fields(1) & "'"
Data6.Refresh
Else
MsgBox (DBCombo1.Text + "已登账")
End If
Data10.Recordset.MoveNext
DBCombo1.Text = Data10.Recordset.Fields(1)
Loop
End Sub


Private Sub DZ4()
Data10.RecordSource = "SELECT CLSCCB.日期,CLSCCB.凭证号,CLSCCB.审核确认,CLSCCB.记账标记 FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND GROUP BY CLSCCB.日期,CLSCCB.凭证号,CLSCCB.审核确认,CLSCCB.记账标记 ORDER BY CLSCCB.日期,VAL(MID(CLSCCB.凭证号,3))"
Data10.Refresh
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
DBCombo1.Text = Data10.Recordset.Fields(1)
Do While Not Data10.Recordset.EOF
Data9.Recordset.FindFirst "凭证号='" & DBCombo1.Text & "'"
If Data9.Recordset.NoMatch Then
Call Chk4
MsgBox (Combo1.Text + DBCombo1.Text + "登账成功")
Data1.Database.Execute "UPDATE CLZZPZ set 记账标记='记' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND CLZZPZ.凭证号='" & Data10.Recordset.Fields(1) & "'"
Data6.Refresh
Else
MsgBox (DBCombo1.Text + "已登账")
End If
Data10.Recordset.MoveNext
DBCombo1.Text = Data10.Recordset.Fields(1)
Loop
End Sub


Private Sub PZDY(PZLB As String, DH As String) ''''无标题

        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军染整软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：

'Select Case Mid(DH, 1, 1)
'       Case "4"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYyf.xls")
'        Excelapp.Sheets(1).Activate
'       Case "2"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYxf.xls")
'        Excelapp.Sheets(1).Activate
'       Case "5"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYzz.xls")
'        Excelapp.Sheets(1).Activate
'       Case "3"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYys.xls")
'        Excelapp.Sheets(1).Activate
'       Case "1"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYxs.xls")
'        Excelapp.Sheets(1).Activate
'End Select

        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDY.xls")
        Excelapp.Sheets(1).Activate


Data1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 6) = Data1.Recordset.Fields(12)
        Excelapp.ActiveSheet.Cells(2, 5) = PZLB
        Excelapp.ActiveSheet.Cells(3, 10) = Trim(DH)
        Excelapp.ActiveSheet.Cells(16, 2) = Data1.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(16, 7) = Data1.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(16, 10) = Data1.Recordset.Fields(8)
i = 5
Do While Not Data1.Recordset.EOF

        Excelapp.ActiveSheet.Cells(i, 1) = Data1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = Data1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 5) = Data1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 8) = Val(Data1.Recordset.Fields(6))
        
        i = i + 1
        
        Excelapp.ActiveSheet.Cells(i, 1) = Data1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = Data1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 5) = Data1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 10) = Val(Data1.Recordset.Fields(6))
        
i = i + 1
Data1.Recordset.MoveNext
Loop

Excelapp.ActiveWindow.Zoom = 100


        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing

End Sub


