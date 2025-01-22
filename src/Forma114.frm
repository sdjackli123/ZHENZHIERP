VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma114 
   BackColor       =   &H00C0E0FF&
   Caption         =   "订单信息"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   1440
      Top             =   9480
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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3720
      TabIndex        =   31
      Text            =   "Text3"
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "订单查询条件"
      Height          =   2175
      Left            =   11640
      TabIndex        =   3
      Top             =   240
      Width           =   6015
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Text            =   "Text4"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "单号"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色别"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "布类"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "欠排"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "备注"
         Height          =   255
         Index           =   6
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "款号"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "多排"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   4
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "刷新"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "退出"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text9"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Text            =   "Text9"
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Text            =   "Text9"
      Top             =   960
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   360
      Top             =   9480
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
      Left            =   480
      Top             =   9480
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
      Left            =   720
      Top             =   9480
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma114.frx":0000
      Height          =   5415
      Left            =   360
      TabIndex        =   15
      Top             =   3960
      Width           =   17415
      _cx             =   30718
      _cy             =   9551
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
      Bindings        =   "Forma114.frx":0015
      Height          =   330
      Left            =   4320
      TabIndex        =   16
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "简称"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   6720
      TabIndex        =   17
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   6720
      TabIndex        =   18
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330366977
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   330366977
      CurrentDate     =   39177
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma114.frx":002A
      Height          =   2415
      Left            =   360
      TabIndex        =   21
      Top             =   1560
      Width           =   11055
      _cx             =   19500
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
      Bindings        =   "Forma114.frx":003F
      Height          =   975
      Left            =   12240
      TabIndex        =   22
      Top             =   3000
      Width           =   5535
      _cx             =   9763
      _cy             =   1720
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
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "色别"
      Height          =   375
      Left            =   6240
      TabIndex        =   30
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "单号"
      Height          =   375
      Index           =   13
      Left            =   3240
      TabIndex        =   29
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "品名"
      Height          =   375
      Index           =   6
      Left            =   6240
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "客户名称"
      Height          =   375
      Index           =   8
      Left            =   3240
      TabIndex        =   27
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "款号"
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   24
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "备注"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   23
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Forma114"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
DTPicker1.value = Date
DTPicker2.value = Date
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = 1
Text9 = ""
DataCombo1 = ""
DataCombo4(1) = ""
DataCombo4(6) = ""
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(9) = 2000
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
On Error Resume Next
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & DataCombo4(1).Text & "'+'%' and "
End If


If Check2(1).value = 1 Then
sql1 = sql1 + "客户 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "色别 like '%'+'" & DataCombo4(6) & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "单号 like '%'+'" & Text9.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "cast(CONVERT(varchar,日期, 23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "((匹数>0 and 欠匹>0) or (计划>0 and 欠排>cast('" & Text4 & "' as real))) and 开卡='N' And "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "备注 like '%'+'" & Text2.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "款号 like '%'+'" & Text1.Text & "'+'%' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "((匹数>0 and 欠匹<0) or (计划>0 and 欠排<cast('" & Text4 & "' as real))) And "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 序号,客户,单号,款号 as 客户单号,品名,色别,色名 as 色号,幅宽,克重,备注,总备注,匹数,计划,总匹,累计,欠匹,欠排,意见 as 印花,技要 as 染色要求,日期,交期,isnull(缸号,'') as 缸号 from sczykpd where (" + sql1 + ")   order by 日期,客户,单号,序号"
Adodc2.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 客户,品名,色别,round(sum(欠匹),1) as 总欠匹,round(sum(欠排),1) as 总欠排,日期 from sczykpd where CONVERT(varchar,日期, 23) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 欠排>0   group by 日期,客户,品名,色别,花型 order by 日期,客户,色别,品名,花型"
Adodc1.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select round(sum(欠匹),1) as 总欠匹,round(sum(欠排),1) as 总欠排 from sczykpd where (" + sql1 + ") "
Adodc3.Refresh

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 600
Next
End If
End Sub

Private Sub Text3_Change()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text3 & "'+'%'  group by 简称"
Adodc4.Refresh
End Sub

Private Sub VSFlexGrid2_Click()
If Adodc2.Recordset.EOF Then
Exit Sub
End If
rs = VSFlexGrid2.Row
cl = VSFlexGrid2.col
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
If cl = 3 Then
Text9.Text = Adodc2.Recordset.Fields(2)
End If
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc2.Recordset.EOF Then
Exit Sub
End If
rs = VSFlexGrid2.Row
cl = VSFlexGrid2.col
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move rs - 1
If cl = 1 Then
Forma11.Text15.Text = Adodc2.Recordset.Fields(21)
Forma11.DataCombo1.Text = Adodc2.Recordset.Fields(1)
Forma11.DataCombo8.Text = Adodc2.Recordset.Fields(2)
Forma11.Text9.Text = Adodc2.Recordset.Fields(3)
Forma11.DataCombo4(1).Text = Adodc2.Recordset.Fields(4)
Forma11.DataCombo4(3).Text = Adodc2.Recordset.Fields(7)
Forma11.DataCombo4(4).Text = Adodc2.Recordset.Fields(15)
Forma11.DataCombo4(5).Text = Adodc2.Recordset.Fields(16)
Forma11.DataCombo4(6).Text = Adodc2.Recordset.Fields(5)
Forma11.Text14 = Adodc2.Recordset.Fields(17)   '印花信息
Forma11.DataCombo4(7) = Adodc2.Recordset.Fields(18)   '染色要求
Forma11.DataCombo4(8).Text = Adodc2.Recordset.Fields(8)
Forma11.Text3.Text = Adodc2.Recordset.Fields(0)
Forma11.Text1.Text = Adodc2.Recordset.Fields(6)
Me.Hide
End If

If cl = 16 Then
Forma172.DataCombo4 = Adodc2.Recordset.Fields(2)
Forma172.DataCombo5 = Adodc2.Recordset.Fields(3)
Forma172.Check2(2).value = 1
Forma172.Show
End If

End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid3.Row
Adodc1.Recordset.Move rs - 1
DataCombo1 = Adodc1.Recordset.Fields(0)
DataCombo4(1) = Adodc1.Recordset.Fields(1)
DataCombo4(6) = Adodc1.Recordset.Fields(2)
Text1 = Adodc1.Recordset.Fields(3)
End Sub

