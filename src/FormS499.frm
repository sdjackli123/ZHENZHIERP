VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FormS499 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ⱦɫ�����ձ���"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   975
      Left            =   15360
      TabIndex        =   83
      Top             =   2160
      Width           =   1815
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   85
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   29
      Left            =   12000
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   28
      Left            =   12000
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   27
      Left            =   12000
      TabIndex        =   79
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   2895
      Index           =   26
      Left            =   13920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   76
      Text            =   "FormS499.frx":0000
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   25
      Left            =   12000
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   24
      Left            =   12000
      TabIndex        =   69
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   23
      Left            =   12000
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   960
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   18
      Left            =   8880
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   22
      Left            =   8880
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   21
      Left            =   8880
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   20
      Left            =   8880
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   19
      Left            =   8880
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   17
      Left            =   8880
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   16
      Left            =   5880
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   15
      Left            =   5880
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   14
      Left            =   5880
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   13
      Left            =   5880
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   12
      Left            =   5880
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   11
      Left            =   5880
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   10
      Left            =   3360
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   4
      Left            =   3360
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   7
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӱ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "FormS499.frx":0006
      Left            =   960
      List            =   "FormS499.frx":0013
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   9
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸ı���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "FormS499.frx":0029
      Height          =   3975
      Left            =   360
      TabIndex        =   1
      Top             =   5040
      Width           =   16575
      _cx             =   29236
      _cy             =   7011
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   15360
      TabIndex        =   4
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329449473
      CurrentDate     =   40055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329449473
      CurrentDate     =   40055
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   15360
      TabIndex        =   20
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      Format          =   329449473
      CurrentDate     =   40055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8400
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8400
      Top             =   10560
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
      Caption         =   "Adodc2"
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
      Height          =   375
      Left            =   8400
      Top             =   10320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8640
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8880
      Top             =   10200
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
      Caption         =   "Adodc5"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8880
      Top             =   10320
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   8880
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
      Caption         =   "Adodc7"
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   8280
      Top             =   10440
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
      Caption         =   "Adodc8"
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
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   8040
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
      Caption         =   "Adodc9"
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
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   7800
      Top             =   10320
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
      Caption         =   "Adodc10"
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   7560
      Top             =   10320
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
      Caption         =   "Adodc11"
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
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   7440
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
      Caption         =   "Adodc12"
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
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   7800
      Top             =   10440
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
      Caption         =   "Adodc13"
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
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   375
      Left            =   9120
      Top             =   10320
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
      Caption         =   "Adodc14"
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
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   9360
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
      Caption         =   "Adodc15"
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormS499.frx":003E
      Height          =   360
      Left            =   1320
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "��̨���"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "FormS499.frx":0053
      Height          =   855
      Left            =   360
      TabIndex        =   39
      Top             =   9000
      Width           =   11535
      _cx             =   20346
      _cy             =   1508
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "FormS499.frx":0068
      Height          =   450
      Left            =   8880
      TabIndex        =   64
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "FormS499.frx":007D
      Height          =   450
      Left            =   8880
      TabIndex        =   65
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "FormS499.frx":0093
      Height          =   450
      Left            =   8880
      TabIndex        =   66
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "FormS499.frx":00A9
      Height          =   450
      Left            =   12000
      TabIndex        =   71
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "Ⱦɫ����"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   11040
      TabIndex        =   82
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   11040
      TabIndex        =   78
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�᲼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   11040
      TabIndex        =   77
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ճ�ʱ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   11040
      TabIndex        =   75
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   11040
      TabIndex        =   74
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ⱦɫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   11040
      TabIndex        =   73
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����ԭ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   13920
      TabIndex        =   72
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ܼ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   7920
      TabIndex        =   63
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƽ��ˮ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   7920
      TabIndex        =   62
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ˮ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   7920
      TabIndex        =   61
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ȾɫŨ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   7920
      TabIndex        =   58
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ⱦɫ��ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   7920
      TabIndex        =   55
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ⱦɫ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   7920
      TabIndex        =   54
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ɫʦ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   4920
      TabIndex        =   51
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "��ɫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   4920
      TabIndex        =   49
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "һ�γɹ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   4920
      TabIndex        =   47
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   4920
      TabIndex        =   45
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   43
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4920
      TabIndex        =   41
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   38
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ɫ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2880
      TabIndex        =   35
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ʒ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   34
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ͻ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   32
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��̨"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   360
      TabIndex        =   31
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƥ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2880
      TabIndex        =   30
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   2880
      TabIndex        =   29
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   360
      TabIndex        =   28
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   360
      TabIndex        =   27
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   2520
      TabIndex        =   26
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2880
      TabIndex        =   25
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ע"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   2880
      TabIndex        =   24
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   15360
      TabIndex        =   23
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   15360
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "FormS499"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_Change()
Text1(8) = Combo1
End Sub

Private Sub Combo1_Click()
Text1(8) = Combo1

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Command1_Click()
'On Error Resume Next

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "rsrbb('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Text1(2).Text & "','" & Text1(3).Text & "','" & Text1(4).Text & "','" & Text1(5).Text & "','" & Text1(6).Text & "','" & Text1(7).Text & "','" & Text1(8).Text & "','" & Text1(9).Text & "','" & Text1(10).Text & "','" & Text1(11).Text & "','" & Text1(12).Text & "','" & Text1(13).Text & "','" & Text1(14).Text & "','" & Text1(15).Text & "','" & Text1(16).Text & "','" & Text1(17).Text & "','" & Text1(18).Text & "','" & Text1(19).Text & "','" & Text1(20).Text & "','" & Text1(21).Text & "','" & Text1(22).Text & "','" & Text1(23).Text & "','" & Text1(24).Text & "','" & Text1(25).Text & "','" & Text1(26).Text & "','" & Text1(27).Text & "','" & Text1(28).Text & "','" & Text1(29).Text & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Adodc1.Refresh

Text1(0).Text = ""
Adodc3.RecordSource = "select max(isnull(���,0)) FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(10).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(10).Text = "1"
Else
Text1(10).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(7).SetFocus
End Sub


Private Sub Command2_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("ȷ���޸���", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 29
Adodc1.Recordset.Fields(i) = Text1(i)
Next
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("ȷ��ɾ����?", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1(8).Text = ""
Text1(0).Text = ""
Adodc3.RecordSource = "select max(isnull(���,0)) FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(10).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(10).Text = "1"
Else
Text1(10).Text = Adodc3.Recordset.Fields(0) + 1
End If
Text1(7).SetFocus
Command1.Enabled = True

Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
Adodc1.RecordSource = "select * FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime) and ���='" & Text1(8) & "' ORDER BY ��� DESC"
Adodc1.Refresh
Adodc3.RecordSource = "select max(isnull(���,0)) FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime) and ���='" & Text1(8) & "'"
Adodc3.Refresh
Text1(10).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(10).Text = "1"
Else
Text1(10).Text = Adodc3.Recordset.Fields(0) + 1
End If
Command1.Enabled = True
Command3.Enabled = False
End Sub


Private Sub Command6_Click()
sql1 = ""
If Check2(0).value = 1 Then
sql1 = sql1 + "�׺�='" & Text1(2) & "' and "
End If
If Check2(3).value = 1 Then
sql1 = sql1 + "��� like '%'+'" & Text1(8) & "'+'%' and "
End If
If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker2.value), "yyyy-MM-dd")
t2 = Format(Trim(DTPicker3.value), "yyyy-MM-dd")
sql1 = sql1 + "CONVERT(varchar(120),����,23) between '" & t1 & "' and '" & t2 & "' and "
End If

If sql1 = "" Then
MsgBox ("��ѡ���ѯ����")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)
Adodc1.RecordSource = "select * FROM rsbb where (" + sql1 + ") order by ����"
Adodc1.Refresh
Adodc7.RecordSource = "select round(sum(ƥ��),1) as �ϼ�ƥ��,round(sum(����),2) as �ϼ����� FROM rsbb where (" + sql1 + ")"
Adodc7.Refresh
End Sub

Private Sub Command8_Click()
Call BBDY(VSFlexGrid1, 6, 7, "Ⱦɫ����")
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
Text1(0).Text = DataCombo1.Text
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT ����,�᲼ FROM ct where ��̨���='" & DataCombo1 & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
Text1(27) = Adodc8.Recordset.Fields(0)
Text1(28) = Adodc8.Recordset.Fields(1)
End If
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Text1(0).Text = DataCombo1.Text
End Sub

Private Sub dataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo2_Change()
Text1(17) = DataCombo2
End Sub

Private Sub DataCombo2_Click(Area As Integer)
Text1(17) = DataCombo2
End Sub

Private Sub DataCombo3_Change()
Text1(18) = DataCombo3
End Sub

Private Sub DataCombo3_Click(Area As Integer)
Text1(18) = DataCombo3
End Sub

Private Sub DataCombo4_Change()
Text1(19) = DataCombo4
End Sub

Private Sub DataCombo4_Click(Area As Integer)
Text1(19) = DataCombo4
End Sub

Private Sub DataCombo5_Change()
Text1(23) = DataCombo5
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT ����ʱ�� FROM rsgysc where Ⱦɫ����='" & DataCombo5 & "'"
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
Text1(24) = Adodc13.Recordset.Fields(0)
End If
End Sub

Private Sub DataCombo5_Click(Area As Integer)
Text1(23) = DataCombo5
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT ����ʱ�� FROM rsgysc where Ⱦɫ����='" & DataCombo5 & "'"
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
Text1(24) = Adodc13.Recordset.Fields(0)
End If
End Sub

Private Sub DTPicker1_Change()
Text1(9).Text = DTPicker1.value
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime) and  "
Adodc1.Refresh
End Sub

Private Sub DTPicker1_CloseUp()
Text1(9).Text = DTPicker1.value
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime) and  "
Adodc1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next

For i = 0 To 29
Text1(i).Text = ""
Next
Text1(9).Text = Date
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
Combo1.Text = ""
Text2 = ""
DTPicker1.value = Date
DTPicker2.value = Date
DTPicker3.value = Date
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * FROM rsbb where cast(CONVERT(varchar,����, 23) as datetime) between cast('" & DTPicker2.value & "' as datetime) and cast('" & DTPicker3.value & "' as datetime) and ��� like '%'+'" & Text1(8).Text & "'+'%' order by ����"
Adodc1.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select max(isnull(���,0)) FROM rsbb where ����=cast('" & DTPicker1.value & "' as datetime)"
Adodc3.Refresh
Text1(10).Text = "1"
If Adodc3.Recordset.EOF Then
Text1(10).Text = "1"
Else
Text1(10).Text = Adodc3.Recordset.Fields(0) + 1
End If


Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT distinct ��̨��� FROM ct "
Adodc5.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT distinct mc FROM rsfs"
Adodc9.Refresh
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT distinct mc FROM rslb"
Adodc10.Refresh
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "SELECT distinct mc FROM rsnd"
Adodc12.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.RecordSource = "SELECT distinct Ⱦɫ���� FROM rsgysc"
Adodc15.Refresh

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 800
VSFlexGrid1.ColWidth(4) = 800

Command1.Enabled = True

Command3.Enabled = False
End Sub

Private Sub Label2_Click()
Forms546.Text1(0) = "Ⱦɫ"
Forms546.Show
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 4
Text1(11) = Now
       Case 8
Text1(12) = Now
       Case 16
Text1(14) = "��"
       Case 18
If Val(Text1(15)) > 3 Then
Text1(15) = 1
Else
Text1(15) = Val(Text1(15)) + 1
End If
End Select
End Sub

Private Sub Label1_DblClick(Index As Integer)
Select Case Index
       Case 16
Text1(14) = "��"
End Select
End Sub

Private Sub Label3_Click()
Adodc6.RecordSource = "select ����,�ͻ�����,'' as ����,ɫ��,Ʒ�� from v_kpd_cx where ����='" & Text1(2).Text & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
Text1(1).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text1(5).Text = ""
Text1(6).Text = ""
Text1(7).Text = ""
Else
Adodc11.RecordSource = "select round(SUM(ƥ��),1) ,round(SUM(����),2) from v_kpd_cx where ����='" & Text1(2).Text & "'"
Adodc11.Refresh
Text1(1).Text = Adodc6.Recordset.Fields(1)
Text1(3).Text = Adodc6.Recordset.Fields(2)
Text1(4).Text = Adodc6.Recordset.Fields(3)
Text1(5).Text = Adodc11.Recordset.Fields(0)
Text1(6).Text = Adodc11.Recordset.Fields(1)
Text1(7).Text = Adodc6.Recordset.Fields(4)
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
       Case 11
       If Not IsDate(Text1(11)) Then
       Text1(11).SetFocus
       End If
       Case 12
       If Not IsDate(Text1(12)) Then
       Text1(12).SetFocus
       End If
End Select
End Sub

Private Sub Text2_Change()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select distinct ��̨��� from ct where ��̨���  like '%'+'" & Text2 & "'+'%' order by ��̨���"
Adodc5.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 2

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select ����,�ͻ�����,'' as ����,ɫ��,Ʒ�� from v_kpd_cx where ����='" & Text1(2).Text & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
'Text1(10).Text = ""
Label3.Enabled = False
Else
Label3.Enabled = True
End If

If InStr(Text1(2).Text, "J") > 0 Then
Text1(2).Text = Mid(Text1(2).Text, 1, Len(Text1(2).Text) - 1)
Call Label3_Click
Text1(6).SetFocus
End If
Case 11
Text1(13) = DateDiff("h", CDate(Text1(11)), CDate(Text1(12)))
Case 12
Text1(13) = DateDiff("h", CDate(Text1(11)), CDate(Text1(12)))
Case 13, 24
Text1(25) = Val(Text1(13)) - Val(Text1(24))
Case 20
Text1(22) = Format(Val(Text1(20)) * Val(Text1(21)), "#0")
Case 21
Text1(22) = Format(Val(Text1(20)) * Val(Text1(21)), "#0")
Case 27, 28
Text1(29) = Val(Text1(27)) + Val(Text1(28))
End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub






