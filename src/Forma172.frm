VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma172 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�ƻ���ѯ"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   14520
      Top             =   11280
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   495
      Left            =   12000
      Top             =   11520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   375
      Left            =   11760
      Top             =   11040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�Ÿ׿���ӡ"
      Height          =   495
      Left            =   18000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���̿���ӡ"
      Height          =   495
      Left            =   18000
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   1200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   8160
      Top             =   10800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ģ���ӡ"
      Height          =   495
      Left            =   12720
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫѡ"
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   12720
      TabIndex        =   65
      Text            =   "Text5"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11400
      TabIndex        =   64
      Text            =   "Text4"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
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
      Left            =   21120
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   8160
      Style           =   1  'Simple Combo
      TabIndex        =   58
      Text            =   "Combo1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11400
      TabIndex        =   54
      Text            =   "Text3"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   14640
      TabIndex        =   47
      Top             =   -120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   15600
      Top             =   0
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Forma172.frx":0000
      Left            =   7680
      List            =   "Forma172.frx":000A
      TabIndex        =   42
      Text            =   "Combo1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   30
      Text            =   "Text8"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   28
      Text            =   "Text2"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   27
      Text            =   "Text8"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   26
      Text            =   "Text8"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command4 
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
      Left            =   19080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
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
      Left            =   19080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
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
      Height          =   375
      Left            =   19080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ѯ����"
      Height          =   1935
      Left            =   14040
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ⱦ��"
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   61
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "˾��"
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   59
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��ͷ"
         Height          =   255
         Index           =   13
         Left            =   1080
         TabIndex        =   57
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ҫ��"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "״̬"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   51
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ɫ��"
         Height          =   255
         Index           =   10
         Left            =   2040
         TabIndex        =   50
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   40
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "��ɫ"
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
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�ͻ�"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4560
      Top             =   10440
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4440
      Top             =   10200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4680
      Top             =   10200
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma172.frx":0016
      Height          =   6975
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   19815
      _cx             =   34951
      _cy             =   12303
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
      ExplorerBar     =   3
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   495
         Left            =   10800
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   330
      Left            =   4800
      TabIndex        =   8
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "Forma172.frx":002B
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "Ʒ��"
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma172.frx":0040
      Height          =   330
      Left            =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���"
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   329580547
      CurrentDate     =   39961
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   329580547
      CurrentDate     =   39961
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Left            =   4800
      TabIndex        =   16
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   7320
      TabIndex        =   17
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma172.frx":0055
      Height          =   1215
      Left            =   360
      TabIndex        =   36
      Top             =   9240
      Width           =   19815
      _cx             =   34951
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
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
      FormatString    =   $"Forma172.frx":006A
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
   Begin MSDataListLib.DataCombo DataCombo6 
      Height          =   330
      Left            =   7320
      TabIndex        =   37
      Top             =   1680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo7 
      Height          =   330
      Left            =   11040
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo8 
      Height          =   330
      Left            =   9240
      TabIndex        =   49
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo9 
      Bindings        =   "Forma172.frx":0141
      Height          =   330
      Left            =   9240
      TabIndex        =   52
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "zt"
      Text            =   "DataCombo3"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "˾��"
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
      Index           =   13
      Left            =   12720
      TabIndex        =   63
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ͷ"
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
      Index           =   11
      Left            =   11400
      TabIndex        =   62
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����״̬"
      Height          =   375
      Left            =   9240
      TabIndex        =   56
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ȾɫҪ��"
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
      Index           =   12
      Left            =   11400
      TabIndex        =   53
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ɫ��"
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
      Index           =   10
      Left            =   9240
      TabIndex        =   48
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   9
      Left            =   11040
      TabIndex        =   45
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Left            =   5520
      TabIndex        =   41
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Left            =   7320
      TabIndex        =   38
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   35
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   34
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   33
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Index           =   4
      Left            =   3840
      TabIndex        =   32
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
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
      Index           =   6
      Left            =   360
      TabIndex        =   24
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
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
      Index           =   5
      Left            =   360
      TabIndex        =   23
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ͻ�����"
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
      Left            =   360
      TabIndex        =   22
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
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
      Left            =   360
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ɫ"
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
      Left            =   4080
      TabIndex        =   20
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   3
      Left            =   4080
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
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
      Index           =   4
      Left            =   7320
      TabIndex        =   18
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Forma172"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public c As Integer
Dim cdbhf As Integer
Dim xzgh As String
Private Declare Function PRINTDLG Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Type PRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Const PD_RETURNDC = &H100
Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_PRINTSETUP = &H40

Private Function ShowPrintDialog(pd As PRINTDLG) As Boolean
    Dim result As Long
    pd.lStructSize = Len(pd)
    pd.hwndOwner = 0
    pd.Flags = PD_RETURNDC Or PD_NOSELECTION Or PD_NOPAGENUMS Or PD_PRINTSETUP
    result = PRINTDLG(pd)
    If result <> 0 Then
        ShowPrintDialog = True
    Else
        ShowPrintDialog = False
    End If
End Function

Private Function GetSelectedPrinter(pd As PRINTDLG) As String
    Dim dm As DEVMODE
    Dim dn As DEVNAMES
    Dim pDevMode As Long
    Dim pDevNames As Long
    Dim PrinterName As String
    
    pDevMode = GlobalLock(pd.hDevMode)
    pDevNames = GlobalLock(pd.hDevNames)
    
    CopyMemory dm, ByVal pDevMode, Len(dm)
    CopyMemory dn, ByVal pDevNames, Len(dn)
    
    PrinterName = StringFromPointer(pDevNames + dn.wDeviceOffset)
    
    GlobalUnlock pd.hDevMode
    GlobalUnlock pd.hDevNames
    
    GetSelectedPrinter = PrinterName
End Function

Private Function StringFromPointer(p As Long) As String
    Dim result As String
    Dim char As Byte
    Do
        CopyMemory char, ByVal p, 1
        If char = 0 Then Exit Do
        result = result & Chr$(char)
        p = p + 1
    Loop
    StringFromPointer = result
End Function






Private Sub Command1_Click()
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "sjkzdbf('')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()
Call bhmx(VSFlexGrid1, 11, 12, DataCombo1.Text)
End Sub

Private Sub Command4_Click()
gyhys = 0
Unload Me
End Sub
Public Sub Command5_Click()
    On Error Resume Next
    Dim sql1 As String
    sql1 = ""

    ' ���ݸ���Check2�ؼ���ֵƴ��SQL��ѯ����
    If Check2(0).value = 1 Then
        sql1 = sql1 + "���� like '%'+'" & DataCombo6.Text & "'+'%' and "
    End If

    If Check2(1).value = 1 Then
        sql1 = sql1 + "�ͻ�����='" & DataCombo1.Text & "' and "
    End If

    If Check2(11).value = 1 Then
        sql1 = sql1 + "����״̬ like '%'+'" & DataCombo9.Text & "'+'%' and "
    End If

    If Check2(2).value = 1 Then
        sql1 = sql1 + "���� like '%'+'" & DataCombo4.Text & "'+'%' and "
    End If

    If Check2(3).value = 1 Then
        sql1 = sql1 + "��� like '%'+'" & DataCombo5.Text & "'+'%' and "
    End If

    If Check2(4).value = 1 Then
        Dim t1 As String
        Dim t2 As String
        t1 = Format(Trim(DTPicker1.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
        t2 = Format(Trim(DTPicker2.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
        sql1 = sql1 + "cast(CONVERT(varchar,����, 120) as datetime) between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
    End If

    If Check2(7).value = 1 Then
        sql1 = sql1 + "Ʒ�� like '%'+'" & DataCombo2.Text & "'+'%' and "
    End If

    If Check2(6).value = 1 Then
        sql1 = sql1 + "�������� like '%'+'" & Combo1.Text & "'+'%' and "
    End If

    If Check2(5).value = 1 Then
        sql1 = sql1 + "ɫ�� like '%'+'" & DataCombo3.Text & "'+'%' and "
    End If

    If Check2(8).value = 1 Then
        sql1 = sql1 + "����='���' and "
    End If

    If Check2(9).value = 1 Then
        sql1 = sql1 + "����='����' and "
    End If

    If Check2(10).value = 1 Then
        sql1 = sql1 + "ɫ�� like '%'+'" & DataCombo8.Text & "'+'%' and "
    End If

    If Check2(12).value = 1 Then
        sql1 = sql1 + "ȾɫҪ�� like '%'+'" & Text3.Text & "'+'%' and "
    End If

    If Check2(13).value = 1 Then
        sql1 = sql1 + "��ͷ = '" & Text4.Text & "' and "
    End If

    If Check2(14).value = 1 Then
        sql1 = sql1 + "˾�� like '%'+'" & Text5.Text & "'+'%' and "
    End If

    If Check2(15).value = 1 Then
        t1 = Format(Trim(DTPicker1.value) + Space(2) + Text2(0) + ":" + Text2(1) + ":" + Text2(2), "yyyy-MM-dd hh:mm:ss")
        t2 = Format(Trim(DTPicker2.value) + Space(2) + Text8(0) + ":" + Text8(1) + ":" + Text8(2), "yyyy-MM-dd hh:mm:ss")
        sql1 = sql1 + "cast(CONVERT(varchar,Ⱦ�׼ƻ�, 120) as datetime) between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and len(Ⱦ�׼ƻ�)>6 and "
    End If

    ' ���û��ѡ���κβ�ѯ����������ʾ���˳�
    If sql1 = "" Then
        MsgBox ("��ѡ���ѯ����")
        Exit Sub
    End If
    ' ȥ������ "and "
    sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

  ' ����Adodc1������Դ��ˢ��
Adodc1.RecordSource = "SELECT * FROM v_kpd_ok WHERE (" + sql1 + ") ORDER BY ����,����"
Adodc1.Refresh
Adodc3.RecordSource = "SELECT count(distinct ����) as �ϼƹ���, " & _
                      "round(sum(CASE WHEN ���� NOT LIKE '%F%' THEN isnull(����, 0) ELSE 0 END), 2) as �ϼ�����, " & _
                      "count(distinct CASE WHEN ���� LIKE '%F%' THEN ���� ELSE NULL END) as ���޹���, " & _
                      "round(sum(CASE WHEN ���� LIKE '%F%' THEN isnull(����, 0) ELSE 0 END), 2) as ��������, " & _
                      "CASE WHEN round(sum(CASE WHEN ���� NOT LIKE '%F%' THEN isnull(����, 0) ELSE 0 END), 2) = 0 " & _
                      "THEN '0.00%' WHEN (sum(CASE WHEN ���� NOT LIKE '%F%' THEN isnull(����, 0) ELSE 0 END) - " & _
                      "sum(CASE WHEN ���� LIKE '%F%' THEN isnull(����, 0) ELSE 0 END)) = 0 " & _
                      "THEN '0.00%' ELSE CAST(ROUND((sum(CASE WHEN ���� LIKE '%F%' THEN isnull(����, 0) ELSE 0 END) / " & _
                      "NULLIF(sum(CASE WHEN ԲͲ���� <> 'N' THEN isnull(����, 0) ELSE 0 END), 0)) * 100, 2) AS varchar) + '%' END as ������, " & _
                      "round(sum(CASE WHEN ԲͲ���� <> 'N' THEN isnull(����, 0) ELSE 0 END), 2) as �ض����� " & _
                      "FROM v_kpd_ok WHERE (" & sql1 & ")"
Adodc3.Refresh



    ' ����VSFlexGrid1���иߺͱ���ɫ
    If VSFlexGrid1.Rows > 1 Then
        Dim i As Integer
        For i = 1 To VSFlexGrid1.Rows - 1
            VSFlexGrid1.RowHeight(i) = 600
            If i / 2 = Int(i / 2) Then
                VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H80000005
            Else
                VSFlexGrid1.Cell(flexcpBackColor, i, 1, i, VSFlexGrid1.Cols - 1) = &H8000000F
            End If
        Next i
    End If

   
    ' ����VSFlexGrid1��VSFlexGrid2���и�ʽ���п�
    VSFlexGrid1.ColFormat(12) = "#0.#"
    VSFlexGrid2.ColFormat(5) = "#0.00%"
    VSFlexGrid2.ColWidth(3) = 1500
    VSFlexGrid2.ColWidth(4) = 1500
    VSFlexGrid2.ColWidth(5) = 1500
    VSFlexGrid2.ColWidth(6) = 1500


    ' �����Զ��庯��gssx
    Call gssx

  
End Sub






Private Sub Command6_Click()
    Dim i As Integer
    With VSFlexGrid1
        For i = 1 To .Rows - 2 ' �ų����һ��
            .Cell(flexcpChecked, i, 3) = True ' ����ѡ��״̬����Ϊѡ��
        Next i
    End With
End Sub
Private Sub Command7_Click()
    Dim i As Integer
    Dim selectedGuoHao As String
    Dim Excelapp As Excel.Application
    Dim wb As Excel.Workbook
    Dim sh As Excel.Worksheet

    ' ����ExcelӦ��ʵ��
    Set Excelapp = New Excel.Application
    If Excelapp Is Nothing Then
        MsgBox "Excel could not be started. Check that your office installation and project references are correct."
        Exit Sub
    End If

    ' ��ģ���ļ�
    Set wb = Excelapp.Workbooks.Open(App.Path & "\��ӡģ��\����\khmb.xls")
    If wb Is Nothing Then
        MsgBox "Template file could not be opened. Check the file path."
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
    End If
    Set sh = wb.Sheets(1)
    sh.Activate

    ' ���� VSFlexGrid1 �е�ѡ���У���ȡ���Ų����� mbdy
    With VSFlexGrid1
        Debug.Print "VSFlexGrid1.Rows: " & .Rows ' ��ӡ������
        For i = 1 To .Rows - 2 ' �ų����һ��
            Debug.Print "Checking row: " & i ' ������Ϣ
            Debug.Print "Checkbox value: " & .Cell(flexcpChecked, i, 3) ' ��ӡ��ѡ���״̬
            If .Cell(flexcpChecked, i, 3) = flexChecked Then ' �����ѡ��ѡ��
                selectedGuoHao = .TextMatrix(i, 4) ' ��ȡ�����еĹ���
                Debug.Print "Selected GuoHao: " & selectedGuoHao ' ��ӡ������Ϣ
                Call mbdy(Adodc5, selectedGuoHao, sh)
            Else
                Debug.Print "Row " & i & " not selected" ' ������Ϣ
            End If
        Next i
    End With

    ' ��ʾExcelӦ��
    Excelapp.Visible = True
    Debug.Print "Excel template should now be visible." ' ��ӡ������Ϣ

    ' �������
    Set sh = Nothing
    Set wb = Nothing
    Set Excelapp = Nothing
End Sub

Private Sub Command8_Click() '''''������ӡ���̿�
    Dim selectedGuoHao As String
    Dim selectedKaHao As String
    Dim pd As PRINTDLG

    ' ��ʾ��ӡ��ѡ�����
    If Not ShowPrintDialog(pd) Then
        Exit Sub ' �û�ȡ���˴�ӡ���˳��ӳ���
    End If

    Dim selectedPrinter As String
    selectedPrinter = GetSelectedPrinter(pd)

    With VSFlexGrid1
        Debug.Print "VSFlexGrid1.Rows: " & .Rows ' ��ӡ������
        For i = 1 To .Rows - 2 ' �ų����һ��
            Debug.Print "Checking row: " & i ' ������Ϣ
            Debug.Print "Checkbox value: " & .Cell(flexcpChecked, i, 3) ' ��ӡ��ѡ���״̬
            If .Cell(flexcpChecked, i, 3) = flexChecked Then ' �����ѡ��ѡ��
                selectedGuoHao = .TextMatrix(i, 4) ' ��ȡ�����еĹ���
                selectedKaHao = .TextMatrix(i, 29) ' ��ȡ��29�еĿ���
                Debug.Print "Selected GuoHao: " & selectedGuoHao ' ��ӡ������Ϣ
                Call lcd22f(Adodc6, Adodc7, selectedGuoHao, selectedKaHao, selectedPrinter)
            Else
                Debug.Print "Row " & i & " not selected" ' ������Ϣ
            End If
        Next i
    End With
End Sub

Private Sub Command9_Click() '''''������ӡ�Ÿ׿�
    Dim selectedGuoHao As String
    Dim pd As PRINTDLG

    ' ��ʾ��ӡ��ѡ�����
    If Not ShowPrintDialog(pd) Then
        Exit Sub ' �û�ȡ���˴�ӡ���˳��ӳ���
    End If

    Dim selectedPrinter As String
    selectedPrinter = GetSelectedPrinter(pd)

    With VSFlexGrid1
        Debug.Print "VSFlexGrid1.Rows: " & .Rows ' ��ӡ������
        For i = 1 To .Rows - 2 ' �ų����һ��
            Debug.Print "Checking row: " & i ' ������Ϣ
            Debug.Print "Checkbox value: " & .Cell(flexcpChecked, i, 3) ' ��ӡ��ѡ���״̬
            If .Cell(flexcpChecked, i, 3) = flexChecked Then ' �����ѡ��ѡ��
                selectedGuoHao = .TextMatrix(i, 4) ' ��ȡ�����еĹ���
                Debug.Print "Selected GuoHao: " & selectedGuoHao ' ��ӡ������Ϣ
                Call pgk(Adodc8, selectedGuoHao, selectedPrinter)
            Else
                Debug.Print "Row " & i & " not selected" ' ������Ϣ
            End If
        Next i
    End With
End Sub



Private Sub DataCombo1_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT Ʒ�� FROM kpd where ���� between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) and �ͻ�����='" & DataCombo1.Text & "' group by Ʒ��"
Adodc3.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT Ʒ�� FROM kpd where ���� between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) and �ͻ�����='" & DataCombo1.Text & "' group by Ʒ��"
Adodc3.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SetDeviceIndependentWindow(Me) '�жϵ�ǰ�ֱ��ʺ����ʱ�ķֱ����Ƿ���ͬ
suiping = Screen.Width / Screen.TwipsPerPixelX  '���㵱ǰ��ˮƽ�ֱ���
cuizhi = Screen.Height / Screen.TwipsPerPixelY '���㵱ǰ�Ĵ�ֱ�ֱ���
If fbl = 1 Then    '��ǰ�ֱ��ʺ����ʱ�ķֱ��ʲ���ͬ
Call ResizeInit(Me)    '����ԭ��������ֵ
Call ResizeForm(Me)    '����������
VSFlexGrid1.FontSize = VSFlexGrid1.FontSize * (suiping / 1366)  ' ��������Ӧ�ĵ���
For i = 0 To 13
Label1(i).FontSize = Label1(i).FontSize * (suiping / 1366)
Next
For i = 1 To 4
Label4(i).FontSize = Label4(i).FontSize * (suiping / 1366)
Next
Label2.FontSize = Label2.FontSize * (suiping / 1366)

DTPicker1.Font.Size = DTPicker1.Font.Size * (suiping / 1366)
DTPicker2.Font.Size = DTPicker2.Font.Size * (suiping / 1366)
Frame2.FontSize = Frame2.FontSize * (suiping / 1366)

For i = 0 To 15
Check2(i).FontSize = Check2(i).FontSize * (suiping / 1366)
Next
Command2.FontSize = Command2.FontSize * (suiping / 1366)
Command3.FontSize = Command3.FontSize * (suiping / 1366)
Command4.FontSize = Command4.FontSize * (suiping / 1366)
Command5.FontSize = Command5.FontSize * (suiping / 1366)

Text1.FontSize = Text1.FontSize * (suiping / 1366)
Text3.FontSize = Text3.FontSize * (suiping / 1366)
Text4.FontSize = Text4.FontSize * (suiping / 1366)
Text5.FontSize = Text5.FontSize * (suiping / 1366)
For i = 0 To 2
Text2(i).FontSize = Text2(i).FontSize * (suiping / 1366)
Text8(i).FontSize = Text8(i).FontSize * (suiping / 1366)
Next
DataCombo1.Font.Size = DataCombo1.Font.Size * (suiping / 1366)
DataCombo2.Font.Size = DataCombo2.Font.Size * (suiping / 1366)
DataCombo3.Font.Size = DataCombo3.Font.Size * (suiping / 1366)
DataCombo5.Font.Size = DataCombo5.Font.Size * (suiping / 1366)
DataCombo6.Font.Size = DataCombo6.Font.Size * (suiping / 1366)
DataCombo7.Font.Size = DataCombo7.Font.Size * (suiping / 1366)
DataCombo8.Font.Size = DataCombo8.Font.Size * (suiping / 1366)
DataCombo9.Font.Size = DataCombo9.Font.Size * (suiping / 1366)
DataCombo10.Font.Size = DataCombo10.Font.Size * (suiping / 1366)
Me.Width = Me.Width * suiping / 1366
Me.Height = Me.Height * cuizhi / 768
End If
DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
DataCombo9.Text = ""
DTPicker1.value = Date
DTPicker2.value = Date
Text1.Text = ""
Combo1 = ""
Text2(0) = "00"
Text2(1) = "00"
Text2(2) = "00"

Text8(0) = "23"
Text8(1) = "59"
Text8(2) = "59"
Text3 = ""
Text4 = ""
Text5 = ""
cdbhf = cdbh
Check2(4).value = 1
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.CommandTimeout = 10000
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_kpd_ok where ���� between cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime)  ORDER BY ����,����,���"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from khZL where ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select zt from v_kpd_zt"
Adodc4.Refresh
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.FrozenCols = 13
VSFlexGrid1.BackColorAlternate = &HCDEEC6
VSFlexGrid1.SelectionMode = flexSelectionListBox
Call gssx

'VSFlexGrid1.SubtotalPosition = flexSTBelow
'VSFlexGrid1.Subtotal flexSTSum, 0, 12, , vbGreen
'VSFlexGrid1.Subtotal flexSTCount, 0, 4, , vbGreen

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
sql1 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
sql2 = "insert into yhcd(�û�,�˵�,���) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
gyhys = 0
sql2 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label2_Click()
Forma170.List3.Clear
Forma170.List3.AddItem "�ƻ�"
Forma170.List3.AddItem "Ⱦ�׼ƻ�"
Forma170.List3.AddItem "Ԥ����Ⱦɫ"
Forma170.List3.AddItem "Ⱦɫ��"
Forma170.List3.AddItem "Ⱦɫ���"
Forma170.List3.AddItem "��ˮ"
Forma170.List3.AddItem "�����ӡ��"
Forma170.List3.AddItem "ĥë"
Forma170.List3.AddItem "���Ͱ�װ"
Forma170.List3.AddItem "�������"
Forma170.Show
End Sub

Private Sub Text1_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ��� from KHZL where ����  like '%'+'" & Text1 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc2.Refresh
End Sub

Private Sub Timer1_Timer()
If Check2(8).value = 1 Then
Call Command5_Click
End If
End Sub

Private Sub VSFlexGrid1_dblClick()
'On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
cl = VSFlexGrid1.col
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
If hysbl = 1 Then
Formh221.DataCombo1(1) = Adodc1.Recordset.Fields(0)
Formh221.DataCombo1(5) = Adodc1.Recordset.Fields(7)
Formh221.DataCombo1(4) = Adodc1.Recordset.Fields(6)
Formh221.DataCombo1(3) = Adodc1.Recordset.Fields(5)
Formh221.Show
hysbl = 0
Me.Hide
End If
If ghcx = 1 And cl = 4 Then
Forma11.Text7 = Adodc1.Recordset.Fields(3)
ghcx = 0
Me.Hide
End If
End Sub

Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C�У���R��
    If c = 22 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        xzgh = .TextMatrix(r, 4)
        Combo1111.Visible = True
        Combo1111.SetFocus
    End If
End With
End Sub
    
        
Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    VSFlexGrid1.Text = Combo1111.Text
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
End If
End Sub

Private Sub Combo1111_LostFocus()
On Error Resume Next
sql1 = "delete from trpd  where ����='" & xzgh & "'"
sql2 = "insert into trpd(����,����) values('" & xzgh & "','" & Combo1111 & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End Sub
Private Sub gssx()
    With VSFlexGrid1
        .FrozenCols = 13
        .BackColorAlternate = &HCDEEC6
        '.SelectionMode = flexSelectionListBox ' ����ѡ��ģʽΪ�б��ģʽ
        .ColWidth(0) = 100 * (suiping / 1366)
        .ColWidth(1) = 800 * (suiping / 1366)
        .ColWidth(2) = 0
        .ColWidth(3) = 900 * (suiping / 1366)
        .ColWidth(4) = 1200 * (suiping / 1366)
        .ColWidth(6) = 600 * (suiping / 1366)
        .ColWidth(5) = 400 * (suiping / 1366)
        .ColWidth(8) = 600 * (suiping / 1366)
        .ColWidth(9) = 600 * (suiping / 1366)
        .ColWidth(10) = 600 * (suiping / 1366)
        .ColWidth(16) = 1000 * (suiping / 1366)
        .ColFormat(12) = "#0.#"

        .TextMatrix(0, 0) = "��¼��"

        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 11, , vbWhite
        .Subtotal flexSTSum, -1, 12, , vbWhite
        .Subtotal flexSTCount, -1, 4, , vbWhite

        If .Rows > 2 Then
            .TextMatrix(.Rows - 1, 1) = "�ϼ�"
        End If

        .RowHeight(0) = 400 * (cuizhi / 768)
       Const CheckboxColumnIndex As Integer = 3
If VSFlexGrid1.Rows > 1 Then
    ' ����ͨ�����̺����༭
    VSFlexGrid1.Editable = flexEDKbdMouse
    ' �������еĸ�ѡ��״̬����Ϊ2��ѡ�У�
    VSFlexGrid1.Cell(flexcpChecked, 1, CheckboxColumnIndex, VSFlexGrid1.Rows - 1, CheckboxColumnIndex) = 2
End If

        If .Rows > 1 Then
            .Row = 1
            .GridLinesFixed = 14
        Else
            .GridLinesFixed = 1
        End If
    End With
End Sub
