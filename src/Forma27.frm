VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma27 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���ܱ���"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "ί��"
      Height          =   255
      Left            =   12840
      TabIndex        =   24
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "����"
      Height          =   255
      Left            =   12840
      TabIndex        =   23
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "����"
      Height          =   255
      Left            =   12840
      TabIndex        =   22
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   21
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   1095
      Left            =   14760
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�鲼"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "֯��"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��̨"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ʼ�"
      Height          =   255
      Left            =   12840
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "׼��"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8760
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
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
      TabIndex        =   25
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo10 
      Height          =   330
      Left            =   9600
      TabIndex        =   26
      Top             =   1080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   6840
      TabIndex        =   27
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   6840
      TabIndex        =   28
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   9600
      TabIndex        =   29
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   9600
      TabIndex        =   30
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   10560
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
         Name            =   "����"
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
      TabIndex        =   31
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   257753091
      CurrentDate     =   39961.3333333333
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   32
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   257753091
      CurrentDate     =   39961.3333333333
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma27.frx":0000
      Height          =   6255
      Left            =   480
      TabIndex        =   33
      Top             =   2400
      Width           =   18015
      _cx             =   31776
      _cy             =   11033
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Bindings        =   "Forma27.frx":0015
      Height          =   1215
      Left            =   480
      TabIndex        =   34
      Top             =   8760
      Width           =   9975
      _cx             =   17595
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2760
      Top             =   10560
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
         Name            =   "����"
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
      Bindings        =   "Forma27.frx":002A
      Height          =   1215
      Left            =   12000
      TabIndex        =   35
      Top             =   8760
      Width           =   5415
      _cx             =   9551
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   1560
      TabIndex        =   36
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo4"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�ʼ�"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   49
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   47
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   46
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   45
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   44
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   43
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   41
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��̨"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   40
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "֯��"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   39
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   18
      Left            =   480
      TabIndex        =   38
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��������"
      Height          =   375
      Index           =   19
      Left            =   480
      TabIndex        =   37
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Forma27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
'On Error Resume Next
sql1 = ""

If Option1.value = True Then
If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "CONVERT(varchar,����, 120) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "�ʼ� like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "���� like '%'+'" & DataCombo6.Text & "'+'%' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "��� like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "left(����,1)='" & DataCombo7.Text & "' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "֯�� like '%'+'" & DataCombo10.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "����Ա like '%'+'" & DataCombo3.Text & "'+'%' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "��̨ like '%'+'" & DataCombo8.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT ����,֯��,���,Ʒ��,����,����,���,����Ա,����,ƥ��,֧�� as ����,��������,�ò�,����,�ʼ�,��ע,��̨,���,�ȼ� FROM v_clbb_zjbb where (" + sql1 + ") ORDER BY ����,��̨,����,֯��,cast(ƥ�� as int)"
Adodc1.Refresh

If Check2(6).value = 1 Then
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT ����Ա,round(sum(����),2) as �ϼƲ���,round(sum(�ò�),2) as �ò��� FROM v_clbb_zjbb where (" + sql1 + ") group by ����Ա order by ����Ա"
Adodc2.Refresh
End If

If Check2(7).value = 1 Then
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT ��̨,Ʒ��,round(sum(����),2) as �ϼƲ���,round(sum(�ò�),2) as �ò��� FROM v_clbb_zjbb where (" + sql1 + ") group by ��̨,Ʒ�� order by ��̨,Ʒ��"
Adodc2.Refresh
End If

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(����),2) as �ϼƲ���,round(sum(�ò�),2) as �ò��� FROM v_clbb_zjbb where (" + sql1 + ")"
Adodc3.Refresh

End If



End Sub

Private Sub Command2_Click()
Call jdmx(VSFlexGrid2, "���ܲ���")
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Call jdmx(VSFlexGrid1, "������ϸ")
End Sub

Private Sub Command6_Click()
'On Error Resume Next
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")

If Option1.value = True Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
'    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "hzclbb('" & t1 & "','" & t2 & "','" & yhm & "','����')"          ' ��ʾ�����ĸ��洢����
'    g_Cmd.Execute           ' ִ�д洢����
'g_Cmd.Cancel
conn.CommandTimeout = 10000    ''''����
sql1 = "DELETE FROM zbclbbhz where �û�='" & yhm & "'"
sql2 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,�׼�,����)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,sum(����),0,0,0,0,0,'" & yhm & "',0,sum(֯��*isnull(����,0)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='һ��' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
sql3 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,�׼�,����)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,0,0,0,0,sum(����),0,'" & yhm & "',0,sum(֯��*isnull(����,0)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
sql4 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,�׼�,����)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,0,0,0,0,0,sum(����),'" & yhm & "',0,sum(֯��*-ISNULL(����,0)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
sql5 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,ҹ��,ҹ��)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,0,sum(����),0,0,0,0,'" & yhm & "',0,sum(isnull(����,0)*cast(ҹ֯ as real)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='һ��' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
sql6 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,ҹ��,ҹ��)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,0,0,0,0,sum(����),0,'" & yhm & "',0,sum(isnull(����,0)*cast(ҹ֯ as real)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic


sql1 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�,ҹ��,����)  SELECT ��̨,�ͻ�,����,Ʒ��,����Ա,0,0,0,0,0,sum(����),'" & yhm & "',0,sum(isnull(����,0)*-cast(ҹ֯ as real)) FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,����Ա"
sql2 = "update zbclbbhz set ����=0 where �û�='" & yhm & "' and ���� is null"
sql3 = "update zbclbbhz set ҹ��=0 where �û�='" & yhm & "' and ҹ�� is null"
sql4 = "update zbclbbhz set ����=0 where �û�='" & yhm & "' and ���� is null"
sql5 = "update zbclbbhz set ���ڷ�Χ='1' where �û�='" & yhm & "'"
sql6 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,����,����,ҹ��,�װ�,ҹ��,����,һ��,����,����,�û�) SELECT '','',Ʒ��,'',����,sum(����),sum(����),sum(ҹ��),sum(�װ�),sum(ҹ��),sum(����),sum(һ��),sum(����),sum(����),�û� FROM zbclbbhz where �û�='" & yhm & "' GROUP BY Ʒ��,����,�û�"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic


sql1 = "delete from zbclbbhz where ���ڷ�Χ='1' and �û�='" & yhm & "'"
sql2 = "update zbclbbhz set ����=�װ�+ҹ��+����+����,һ��=�װ�+ҹ�� where �û�='" & yhm & "'"
sql3 = "update zbclbbhz set һ����=round(һ��/����*100,1) where �û�='" & yhm & "'"
sql4 = "update zbclbbhz set �׼�=����/�װ� where �û�='" & yhm & "' and �װ�<>0"
sql5 = "update zbclbbhz set ҹ��=ҹ��/ҹ�� where �û�='" & yhm & "' and ҹ��<>0"
sql6 = "update zbclbbhz set ���ڷ�Χ=CONVERT(varchar,'" & t1 & "', 23)+'/'+CONVERT(varchar,'" & t2 & "', 23) where �û�='" & yhm & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic

Adodc1.RecordSource = "SELECT ����,����,�װ�,ҹ��,һ��,һ����,����,����,���� from zbclbbhz where �û�='" & yhm & "'"
Adodc1.Refresh

Adodc2.RecordSource = "SELECT ����,round(sum(�װ�),2) as �װ�,round(sum(ҹ��),2) as ҹ��,round(sum(һ��),2) as һ��,round(round(sum(һ��),2)/round(sum(����),2)*100,1) һ����,round(sum(����),1) as ����,round(sum(����),1) as ����,round(sum(����),1) as ���� FROM zbclbbhz where �û�='" & yhm & "' group by ���� order by ����"
Adodc2.Refresh

Adodc3.RecordSource = "SELECT round(sum(����),2) as ����,round(sum(һ��),2) as һ�� FROM zbclbbhz where �û�='" & yhm & "'"
Adodc3.Refresh

End If

If Option2.value = True Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
''    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "hzclbb('" & T1 & "','" & T2 & "','" & yhm & "','����')"          ' ��ʾ�����ĸ��洢����
'    g_Cmd.Execute           ' ִ�д洢����
'g_Cmd.Cancel

conn.CommandTimeout = 10000    ''''����
sql1 = "DELETE FROM zbclbbhz where �û�='" & yhm & "'"
sql2 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,sum(isnull(����,0)),0,0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='һ��' and ��̨ not in(select distinct ��� from gys where ip like'%Z%') GROUP BY ��̨,�ͻ�,Ʒ��"
sql3 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,0,sum(isnull(����,0)),0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='����' GROUP BY ��̨,�ͻ�,Ʒ��"
sql4 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,0,0,sum(isnull(����,0)),'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='����' GROUP BY ��̨,�ͻ�,Ʒ��"
sql5 = "update zbclbbhz set ����='' where ���� is null"
sql6 = "update zbclbbhz set �ͻ�='' where �ͻ� is null"
sql7 = "update zbclbbhz set Ʒ��='' where Ʒ�� is null"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic

sql1 = "update zbclbbhz set ���ڷ�Χ='1' where �û�='" & yhm & "'"
sql2 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�) SELECT ����,�ͻ�,Ʒ��,sum(isnull(����,0)),sum(isnull(һ��,0)),sum(isnull(����,0)),sum(isnull(����,0)),�û� FROM zbclbbhz where �û�='" & yhm & "' GROUP BY ����,�ͻ�,Ʒ��,�û�"
sql3 = "delete from zbclbbhz where ���ڷ�Χ='1' and �û�='" & yhm & "'"
sql4 = "insert into zbclbbhz(����,����,һ��,����,����,�û�) SELECT ����+'�ϼ�',sum(isnull(����,0)),sum(isnull(һ��,0)),sum(isnull(����,0)),sum(isnull(����,0)),�û� FROM zbclbbhz where �û�='" & yhm & "' GROUP BY ����,�û�"
sql5 = "update zbclbbhz set ����=һ��+����+���� where �û�='" & yhm & "' and ���� not like '�ϼ�%' and �ͻ� not like '�ϼ�%'"
sql6 = "update zbclbbhz set һ����=round((һ��)/����*100,1),������=round((����)/����*100,1),������=round((����)/����*100,1) where �û�='" & yhm & "' and ���� not like '�ϼ�%' and �ͻ� not like '�ϼ�%' and ����<>0"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic

sql1 = "update zbclbbhz set ����=һ��+����+���� where �û�='" & yhm & "'  and �ͻ� like '%�ϼ�'"
sql2 = "update zbclbbhz set һ����=round((һ��)/����*100,1),������=round((����)/����*100,1),������=round((����)/����*100,1) where �û�='" & yhm & "' and �ͻ� like '%�ϼ�' and ����<>0"
sql3 = "update zbclbbhz set ����=isnull(һ��,0)+isnull(����,0)+isnull(����,0) where �û�='" & yhm & "'  and ���� like '%�ϼ�'"
sql4 = "update zbclbbhz set һ����=round((һ��)/����*100,1),������=round((����)/����*100,1),������=round((����)/����*100,1) where �û�='" & yhm & "' and ���� like '%�ϼ�' and ����<>0"
sql5 = "update zbclbbhz set ���ڷ�Χ=CONVERT(varchar,'" & t1 & "', 23)+'/'+CONVERT(varchar,'" & t2 & "', 23) where �û�='" & yhm & "'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic


Adodc1.RecordSource = "SELECT ����,�ͻ�,Ʒ��,����,һ��,һ����,����,������,����,������ from zbclbbhz where �û�='" & yhm & "' order by ����,�ͻ�,Ʒ��"
Adodc1.Refresh

Adodc2.RecordSource = "SELECT ����,round(sum(����),2) as ����,round(sum(һ��),2) as һ��,round(round(sum(һ��),2)/round(sum(����),2)*100,1) һ����,round(sum(����),2) as ����,round(round(sum(����),2)/round(sum(����),2)*100,1) ������,round(sum(����),2) as ����,round(round(sum(����),2)/round(sum(����),2)*100,1) ������ FROM zbclbbhz where �û�='" & yhm & "' and ���� not like '%�ϼ�%' group by ���� order by ���� desc"
Adodc2.Refresh

Adodc3.RecordSource = "SELECT round(sum(����),2) as ����,round(sum(һ��),2) as һ��,round(sum(����),2) as ����,round(sum(����),2) as ���� FROM zbclbbhz where �û�='" & yhm & "' and ���� not like '%�ϼ�%'"
Adodc3.Refresh

End If

If Option3.value = True Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
'    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "hzclbb('" & t1 & "','" & t2 & "','" & yhm & "','ί��')"          ' ��ʾ�����ĸ��洢����
'    g_Cmd.Execute           ' ִ�д洢����
'g_Cmd.Cancel

conn.CommandTimeout = 10000    ''''����

sql1 = "DELETE FROM zbclbbhz where �û�='" & yhm & "'"
sql2 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,sum(����),0,0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='һ��' and ��̨ in(select distinct ��� from gys where ip like'%Z%') GROUP BY ��̨,�ͻ�,Ʒ��"
sql3 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,0,sum(����),0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='����' GROUP BY ��̨,�ͻ�,Ʒ��"
sql4 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,Ʒ��,0,0,0,sum(����),'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and �ȼ�='����' GROUP BY ��̨,�ͻ�,Ʒ��"
sql5 = "update zbclbbhz set ���ڷ�Χ='1' where �û�='" & yhm & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic


sql1 = "insert into zbclbbhz(����,�ͻ�,Ʒ��,����,һ��,����,����,�û�) SELECT ����,�ͻ�,Ʒ��,sum(����),sum(һ��),sum(����),sum(����),�û� FROM zbclbbhz where �û�='" & yhm & "' GROUP BY ����,�ͻ�,Ʒ��,�û�"
sql2 = "delete from zbclbbhz where ���ڷ�Χ='1' and �û�='" & yhm & "'"
sql3 = "update zbclbbhz set ����=һ��+����+���� where �û�='" & yhm & "'"
sql4 = "update zbclbbhz set һ����=round((һ��)/����*100,1),������=round((����)/����*100,1),������=round((����)/����*100,1) where �û�='" & yhm & "'"
sql5 = "update zbclbbhz set ���ڷ�Χ=CONVERT(varchar,'" & t1 & "', 23)+'/'+CONVERT(varchar,'" & t2 & "', 23) where �û�='" & yhm & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic

Adodc1.RecordSource = "SELECT ����,�ͻ�,Ʒ��,����,һ��,һ����,����,������,����,������ from zbclbbhz where �û�='" & yhm & "'"
Adodc1.Refresh

Adodc2.RecordSource = "SELECT ����,round(sum(����),2) as ����,round(sum(һ��),2) as һ��,round(round(sum(һ��),2)/round(sum(����),2)*100,1) һ����,round(sum(����),2) as ����,round(round(sum(����),2)/round(sum(����),2)*100,1) ������,round(sum(����),2) as ����,round(round(sum(����),2)/round(sum(����),2)*100,1) ������ FROM zbclbbhz where �û�='" & yhm & "' group by ���� order by ���� desc"
Adodc2.Refresh

Adodc3.RecordSource = "SELECT round(sum(����),2) as ����,round(sum(һ��),2) as һ��,round(sum(����),2) as ����,round(sum(����),2) as ���� FROM zbclbbhz where �û�='" & yhm & "'"
Adodc3.Refresh

End If

If Option4.value = True Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
'    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "hzclbb('" & t1 & "','" & t2 & "','" & yhm & "','�ʼ�')"          ' ��ʾ�����ĸ��洢����
'    g_Cmd.Execute           ' ִ�д洢����
'g_Cmd.Cancel
conn.CommandTimeout = 10000    ''''����

sql1 = "DELETE FROM zbclbbhz where �û�='" & yhm & "'"
sql2 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,sum(����),0,0,0,0,0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='һ��' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
sql3 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,0,0,0,0,sum(����),0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
sql4 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,0,0,0,0,0,sum(����),'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='�װ�' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
sql5 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,0,sum(����),0,0,0,0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='һ��' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
sql6 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,0,0,0,0,sum(����),0,'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic


sql1 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�)  SELECT ��̨,�ͻ�,����,Ʒ��,�ʼ�,0,0,0,0,0,sum(����),'" & yhm & "' FROM v_clbb_zjbb WHERE ���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and ���='ҹ��' and �ȼ�='����' GROUP BY ��̨,�ͻ�,����,Ʒ��,�ʼ�"
sql2 = "update zbclbbhz set ���ڷ�Χ='1' where �û�='" & yhm & "'"
sql3 = "insert into zbclbbhz(����,�ͻ�,����,Ʒ��,����,�װ�,ҹ��,����,һ��,����,����,�û�) SELECT '','',����,'',����,sum(�װ�),sum(ҹ��),sum(����),sum(һ��),sum(����),sum(����),�û� FROM zbclbbhz where �û�='" & yhm & "' GROUP BY ����,����,�û�"
sql4 = "delete from zbclbbhz where ���ڷ�Χ='1' and �û�='" & yhm & "'"
sql5 = "update zbclbbhz set ����=�װ�+ҹ��+����+����,һ��=�װ�+ҹ�� where �û�='" & yhm & "'"
sql6 = "update zbclbbhz set һ����=round(һ��/����*100,1) where �û�='" & yhm & "'"
sql7 = "update zbclbbhz set ���ڷ�Χ=CONVERT(varchar,'" & t1 & "', 23)+'/'+CONVERT(varchar,'" & t2 & "', 23) where �û�='" & yhm & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic
RD.Open sql7, conn, adOpenStatic, adLockOptimistic



Adodc1.RecordSource = "SELECT ����,����,�װ�,ҹ��,һ��,һ����,����,����,���� from zbclbbhz where �û�='" & yhm & "'"
Adodc1.Refresh

Adodc2.RecordSource = "SELECT ����,round(sum(�װ�),2) as �װ�,round(sum(ҹ��),2) as ҹ��,round(sum(һ��),2) as һ��,round(round(sum(һ��),2)/round(sum(����),2)*100,1) һ����,round(sum(����),1) as ����,round(sum(����),1) as ����,round(sum(����),1) as ���� FROM zbclbbhz where �û�='" & yhm & "' group by ���� order by ����"
Adodc2.Refresh

Adodc3.RecordSource = "SELECT round(sum(����),2) as ����,round(sum(һ��),2) as һ�� FROM zbclbbhz where �û�='" & yhm & "'"
Adodc3.Refresh

End If

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

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT ����,֯��,��� as ��ͬ��,Ʒ��,����,����,����Ա,����,ƥ��,֧�� as ����,��������,�ò�,����,�ʼ�,��ע,��̨,���,�ȼ� FROM v_clbb_zjbb where ���� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime)  ORDER BY ����,֯��"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid2.ColWidth(0) = 200
End Sub


