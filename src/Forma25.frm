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
      Caption         =   "��ѯ����"
      Height          =   1095
      Left            =   14520
      TabIndex        =   14
      Top             =   360
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "֯��"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�鲼"
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
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
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
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   975
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
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "��ϸ����"
      Height          =   375
      Left            =   12840
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ձ���"
      Height          =   375
      Left            =   12840
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ʼ챨��"
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
      Caption         =   "��������"
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
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   18
      Left            =   480
      TabIndex        =   46
      Top             =   480
      Width           =   1095
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
      TabIndex        =   45
      Top             =   960
      Width           =   495
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
      Index           =   8
      Left            =   6240
      TabIndex        =   44
      Top             =   1560
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
      TabIndex        =   43
      Top             =   960
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
      Top             =   480
      Width           =   615
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
      TabIndex        =   41
      Top             =   480
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
      Index           =   2
      Left            =   9120
      TabIndex        =   40
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   39
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   38
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   37
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   36
      Top             =   1080
      Width           =   255
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
sql1 = sql1 + "���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
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
sql1 = sql1 + "���� like '%'+'" & DataCombo7.Text & "'+'%' and "
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
Adodc1.RecordSource = "SELECT ����,֯��,���,Ʒ��,Ͳ��,����,����,���,����Ա,����,ƥ��,����,֧�� as ����,��������,�ò�,����,�ʼ�,����,��̨ as ����,��� as ��̨,�ȼ� FROM v_clbb_zjbb where (" + sql1 + ")  ORDER BY ����,��̨,����,֯��,cast(ƥ�� as int)"
Adodc1.Refresh

If Check2(6).value = 1 Then


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "hzclbb('" & t1 & "','" & t2 & "','" & yhm & "','����')"          ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel


Adodc2.RecordSource = "SELECT ����,����,�װ�,�׼�,����,ҹ��,ҹ��,ҹ��,һ��,һ����,����,����,����,���� from clbbhz where �û�='" & yhm & "' and ���� like '%'+'" & DataCombo3 & "'+'%' order by ����"
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
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjzjsc('" & t1 & "','" & t2 & "')"          ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
g_Cmd.Cancel


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT �ʼ�,����,round(sum(����),2) as �ϼƲ��� FROM zjbbf where (" + sql1 + ") group by �ʼ�,���� order by �ʼ�,����"
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
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
'    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "cjzjsc('" & DTPicker1.Value & "','" & DTPicker2.Value & "')"          ' ��ʾ�����ĸ��洢����
'    g_Cmd.Execute           ' ִ�д洢����
'g_Cmd.Cancel


Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT ����,����,round(sum(����),2) as �ϼƲ��� FROM zjbbf where (" + sql1 + ") group by ����,���� order by ����,����"
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
Adodc3.RecordSource = "SELECT round(sum(����),2) as �ϼƲ��� FROM v_clbb_zjbb where (" + sql1 + ")"
Adodc3.Refresh

End If

If Option2.value = True Then
If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "���� between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If


If Check2(5).value = 1 Then
sql1 = sql1 + "֯�� like '%'+'" & DataCombo10.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT * from v_clbb_rbb where (" + sql1 + ") ORDER BY ����,֯��"
Adodc1.Refresh


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(�ղ���),2) as �ϼƲ��� FROM v_clbb_rbb where (" + sql1 + ")"
Adodc3.Refresh

End If

If Option3.value = True Then
If Check2(0).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text1(0) + ":" + Text1(1) + ":" + Text1(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text3(0) + ":" + Text3(1) + ":" + Text3(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "���� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If


If Check2(5).value = 1 Then
sql1 = sql1 + "֯�� like '%'+'" & DataCombo10.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "SELECT Ʒ��,sum(�ϼƲ���) as �ܲ���,�ȼ� from v_clbb_rbb1 where (" + sql1 + ") group by Ʒ��,�ȼ� ORDER BY Ʒ��,�ȼ�"
Adodc1.Refresh


Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT round(sum(�ϼƲ���),2) as �ϼƲ��� FROM v_clbb_rbb1 where (" + sql1 + ")"
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

Private Sub Command5_Click()
Call jdmx(VSFlexGrid3, "���ܲ���")
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
Adodc1.RecordSource = "SELECT ����,֯��,��� as ��ͬ��,Ʒ��,Ͳ��,����,����,����Ա,����,ƥ��,����,֧�� as ����,��������,�ò�,����,�ʼ�,����,��̨ as ����,��� as ��̨,�ȼ� FROM v_clbb_zjbb where ���� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime)  ORDER BY ����,֯��"
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

