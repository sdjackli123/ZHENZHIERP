VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms511 
   BackColor       =   &H00C0E0FF&
   Caption         =   "����ɨ��"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����ģʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   225
      Top             =   5040
      Width           =   4215
      Begin VB.OptionButton Option4 
         Caption         =   "�ֶ�"
         Height          =   495
         Left            =   2640
         TabIndex        =   227
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "�Զ�"
         Height          =   495
         Left            =   840
         TabIndex        =   226
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   223
      Text            =   "Forms511.frx":0000
      Top             =   6240
      Width           =   6135
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   4680
      Top             =   10560
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
      Left            =   4800
      Top             =   10680
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
   Begin VB.OptionButton Option2 
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
      Height          =   855
      Left            =   6840
      TabIndex        =   220
      Top             =   4320
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "��̨����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6840
      TabIndex        =   219
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   218
      Text            =   "Text11"
      Top             =   1800
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   4080
      Top             =   10680
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4200
      ItemData        =   "Forms511.frx":0007
      Left            =   12000
      List            =   "Forms511.frx":0009
      Style           =   1  'Checkbox
      TabIndex        =   207
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "����Ա����Ż���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8520
      TabIndex        =   103
      Top             =   2880
      Width           =   6615
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   4320
         TabIndex        =   228
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   224
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   216
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   215
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   1800
         TabIndex        =   211
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   120
         TabIndex        =   210
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   960
         TabIndex        =   209
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "/"
         Height          =   495
         Index           =   13
         Left            =   5760
         TabIndex        =   206
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   5040
         TabIndex        =   116
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   495
         Left            =   5760
         TabIndex        =   115
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   4320
         TabIndex        =   114
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   3480
         TabIndex        =   113
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   3480
         TabIndex        =   112
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   2640
         TabIndex        =   111
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2640
         TabIndex        =   110
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1800
         TabIndex        =   109
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1800
         TabIndex        =   108
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   960
         TabIndex        =   107
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   960
         TabIndex        =   106
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   105
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   104
         Top             =   360
         Width           =   615
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms511.frx":000B
      Height          =   1095
      Left            =   360
      TabIndex        =   62
      Top             =   2400
      Width           =   6255
      _cx             =   11033
      _cy             =   1931
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
      FormatString    =   $"Forms511.frx":0020
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
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   102
      Text            =   "Text10"
      Top             =   1080
      Width           =   2655
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forms511.frx":00F5
      Height          =   1335
      Left            =   360
      TabIndex        =   63
      Top             =   3600
      Width           =   6255
      _cx             =   11033
      _cy             =   2355
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Forms511.frx":010A
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   98
      Text            =   "Text1"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   360
      TabIndex        =   64
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   116
         Left            =   8640
         TabIndex        =   205
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   115
         Left            =   8640
         TabIndex        =   204
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   114
         Left            =   8640
         TabIndex        =   203
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   113
         Left            =   8640
         TabIndex        =   202
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   112
         Left            =   8640
         TabIndex        =   201
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   111
         Left            =   8640
         TabIndex        =   200
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   110
         Left            =   8640
         TabIndex        =   199
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   109
         Left            =   8640
         TabIndex        =   198
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   108
         Left            =   7920
         TabIndex        =   197
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   107
         Left            =   7920
         TabIndex        =   196
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   106
         Left            =   7920
         TabIndex        =   195
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   105
         Left            =   7920
         TabIndex        =   194
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   104
         Left            =   7920
         TabIndex        =   193
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   103
         Left            =   7920
         TabIndex        =   192
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   102
         Left            =   7920
         TabIndex        =   191
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   101
         Left            =   7920
         TabIndex        =   190
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   100
         Left            =   7920
         TabIndex        =   189
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   99
         Left            =   7200
         TabIndex        =   188
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   98
         Left            =   7200
         TabIndex        =   187
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   97
         Left            =   7200
         TabIndex        =   186
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   96
         Left            =   7200
         TabIndex        =   185
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   95
         Left            =   7200
         TabIndex        =   184
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   94
         Left            =   7200
         TabIndex        =   183
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   93
         Left            =   7200
         TabIndex        =   182
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   92
         Left            =   7200
         TabIndex        =   181
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   91
         Left            =   7200
         TabIndex        =   180
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   90
         Left            =   6480
         TabIndex        =   179
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   8520
         X2              =   8520
         Y1              =   120
         Y2              =   5520
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   89
         Left            =   6480
         TabIndex        =   178
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   88
         Left            =   6480
         TabIndex        =   177
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   87
         Left            =   6480
         TabIndex        =   176
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   86
         Left            =   6480
         TabIndex        =   175
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   85
         Left            =   6480
         TabIndex        =   174
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   84
         Left            =   6480
         TabIndex        =   173
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   83
         Left            =   6480
         TabIndex        =   172
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   7800
         X2              =   7800
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   82
         Left            =   6480
         TabIndex        =   171
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   81
         Left            =   5760
         TabIndex        =   170
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   80
         Left            =   5760
         TabIndex        =   169
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   79
         Left            =   5760
         TabIndex        =   168
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   78
         Left            =   5760
         TabIndex        =   167
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   77
         Left            =   5760
         TabIndex        =   166
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   76
         Left            =   5760
         TabIndex        =   165
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   7080
         X2              =   7080
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   75
         Left            =   5760
         TabIndex        =   164
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   74
         Left            =   5760
         TabIndex        =   163
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   73
         Left            =   5760
         TabIndex        =   162
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   72
         Left            =   5040
         TabIndex        =   161
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   71
         Left            =   5040
         TabIndex        =   160
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   70
         Left            =   5040
         TabIndex        =   159
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   69
         Left            =   5040
         TabIndex        =   158
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   6360
         X2              =   6360
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   68
         Left            =   5040
         TabIndex        =   157
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   67
         Left            =   5040
         TabIndex        =   156
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   66
         Left            =   5040
         TabIndex        =   155
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   65
         Left            =   5040
         TabIndex        =   154
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   64
         Left            =   5040
         TabIndex        =   153
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   63
         Left            =   4320
         TabIndex        =   152
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   62
         Left            =   4320
         TabIndex        =   151
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   61
         Left            =   4320
         TabIndex        =   150
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   60
         Left            =   4320
         TabIndex        =   149
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   59
         Left            =   4320
         TabIndex        =   148
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   58
         Left            =   4320
         TabIndex        =   147
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   57
         Left            =   4320
         TabIndex        =   146
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   56
         Left            =   4320
         TabIndex        =   145
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   55
         Left            =   4320
         TabIndex        =   144
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   4920
         X2              =   4920
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   54
         Left            =   3600
         TabIndex        =   143
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   53
         Left            =   3600
         TabIndex        =   142
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   52
         Left            =   3600
         TabIndex        =   141
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   51
         Left            =   3600
         TabIndex        =   140
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   50
         Left            =   3600
         TabIndex        =   139
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   49
         Left            =   3600
         TabIndex        =   138
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   48
         Left            =   3600
         TabIndex        =   137
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   47
         Left            =   3600
         TabIndex        =   136
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   46
         Left            =   3600
         TabIndex        =   135
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   45
         Left            =   2880
         TabIndex        =   134
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   44
         Left            =   2880
         TabIndex        =   133
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   43
         Left            =   2880
         TabIndex        =   132
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   42
         Left            =   2880
         TabIndex        =   131
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   41
         Left            =   2880
         TabIndex        =   130
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   3480
         X2              =   3480
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   40
         Left            =   2880
         TabIndex        =   129
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   2880
         TabIndex        =   128
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   38
         Left            =   2880
         TabIndex        =   127
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   37
         Left            =   2880
         TabIndex        =   126
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   36
         Left            =   2160
         TabIndex        =   125
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   35
         Left            =   2160
         TabIndex        =   124
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   34
         Left            =   2160
         TabIndex        =   123
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   3
         Visible         =   0   'False
         X1              =   2760
         X2              =   2760
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   33
         Left            =   2160
         TabIndex        =   122
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   32
         Left            =   2160
         TabIndex        =   121
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   31
         Left            =   2160
         TabIndex        =   120
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   30
         Left            =   2160
         TabIndex        =   119
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   29
         Left            =   2160
         TabIndex        =   118
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   8640
         TabIndex        =   117
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   600
         X2              =   600
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   92
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   91
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   90
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   89
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   88
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   87
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   86
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   85
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   84
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   720
         TabIndex        =   83
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   720
         TabIndex        =   82
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   720
         TabIndex        =   81
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   720
         TabIndex        =   80
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   720
         TabIndex        =   79
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   1320
         X2              =   1320
         Y1              =   120
         Y2              =   5400
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   720
         TabIndex        =   78
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   720
         TabIndex        =   77
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   720
         TabIndex        =   76
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   720
         TabIndex        =   75
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   1440
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   1440
         TabIndex        =   73
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   21
         Left            =   1440
         TabIndex        =   72
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   1440
         TabIndex        =   71
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   23
         Left            =   1440
         TabIndex        =   70
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   1440
         TabIndex        =   69
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   25
         Left            =   1440
         TabIndex        =   68
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   26
         Left            =   1440
         TabIndex        =   67
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   27
         Left            =   1440
         TabIndex        =   66
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   5520
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   28
         Left            =   2160
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   6240
      Top             =   10680
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
      Height          =   375
      Left            =   4800
      Top             =   10680
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
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10440
      Top             =   0
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8520
      TabIndex        =   1
      Top             =   1320
      Width           =   6615
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   2160
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   3120
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   4080
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   975
   End
   Begin VB.Data Data8 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6960
      Top             =   10680
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
      Left            =   6960
      Top             =   10680
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
      Left            =   6720
      Top             =   10680
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Left            =   7680
      Top             =   10680
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
      Height          =   375
      Left            =   7680
      Top             =   10680
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
      Left            =   6240
      Top             =   10680
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
   Begin VB.Label Label22 
      Caption         =   "����Ա"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   222
      Top             =   5280
      Width           =   6135
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      TabIndex        =   221
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
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
      Left            =   6120
      TabIndex        =   217
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3840
      TabIndex        =   214
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3240
      TabIndex        =   213
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   4440
      TabIndex        =   212
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   208
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ǰ����"
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
      TabIndex        =   101
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ǰ����"
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
      Left            =   3360
      TabIndex        =   99
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   35
      Left            =   11280
      TabIndex        =   97
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   34
      Left            =   11280
      TabIndex        =   96
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   33
      Left            =   11280
      TabIndex        =   95
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   32
      Left            =   11280
      TabIndex        =   94
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ˮϴ"
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
      Index           =   0
      Left            =   13560
      TabIndex        =   93
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2880
      TabIndex        =   61
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�׺�"
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
      TabIndex        =   60
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�Ų���"
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
      Left            =   8520
      TabIndex        =   59
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�ѱ���"
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
      Left            =   8520
      TabIndex        =   58
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "δ����"
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
      Left            =   12120
      TabIndex        =   57
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�����"
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
      Left            =   12120
      TabIndex        =   56
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6960
      TabIndex        =   55
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6960
      TabIndex        =   54
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   6960
      TabIndex        =   53
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   6960
      TabIndex        =   52
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   6960
      TabIndex        =   51
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7680
      TabIndex        =   50
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   7680
      TabIndex        =   49
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   7680
      TabIndex        =   48
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   7680
      TabIndex        =   47
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   7680
      TabIndex        =   46
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   8400
      TabIndex        =   45
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   8400
      TabIndex        =   44
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   8400
      TabIndex        =   43
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   8400
      TabIndex        =   42
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   8400
      TabIndex        =   41
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   9120
      TabIndex        =   40
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   9120
      TabIndex        =   39
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   9120
      TabIndex        =   38
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   9120
      TabIndex        =   37
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   9120
      TabIndex        =   36
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   9840
      TabIndex        =   35
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   9840
      TabIndex        =   34
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   9840
      TabIndex        =   33
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   9840
      TabIndex        =   32
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   25
      Left            =   9840
      TabIndex        =   31
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   26
      Left            =   10560
      TabIndex        =   30
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   27
      Left            =   10560
      TabIndex        =   29
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   28
      Left            =   10560
      TabIndex        =   28
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   29
      Left            =   10560
      TabIndex        =   27
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   30
      Left            =   10560
      TabIndex        =   26
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "ˮϴ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   31
      Left            =   11280
      TabIndex        =   25
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ǰ�׺�"
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
      TabIndex        =   24
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4560
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "Forms511"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l1, L2, hgxx As String
Dim gybb, gybh As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim cdbhf, sjsx As Integer
Dim dqgx As Integer
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End Sub


Private Sub Command2_Click()
On Error Resume Next
If Text9.Text = "" Then
MsgBox ("���������Ա��Ϣ��")
Exit Sub
End If

If Adodc1.Recordset.EOF Then
Label11.Caption = "��ɨ�����̿�"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
Else

bz = ""       '����
yg = Text9.Text          'Ա��

If Text3.Text <> "" And Len(Text3.Text) = 2 Then

If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then       ''''''''''''''''''�Ų�
Adodc3.RecordSource = "select * from PBCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmpb('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then    '''''''''''''''''Ⱦɫ
Adodc3.RecordSource = "select * from RSCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmrs('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If



If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then    '''''''''''''''''��ˮ
Adodc3.RecordSource = "select * from TSCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmts('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then    '''''''''''''''''���
Adodc3.RecordSource = "select * from HGCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmhg('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then    '''''''''''''''''С����
Adodc3.RecordSource = "select * from XDCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmxd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then    '''''''''''''''''����
Adodc3.RecordSource = "select * from DDCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmdd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & Date & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "')"       ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
End If

Else
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If


Label11.Caption = "��ɨ�����̿�"
Text2.Text = ""
Exit Sub
End If


End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Command4_Click()
  Dim text12Value As String
    Dim parts() As String
    Dim result As String

    ' ���� Text12 �ǵ�ǰ���е�һ���ı���ؼ�
    text12Value = Text12.Text  ' ʹ�ÿؼ��� Text ���Ի�ȡ��ֵ

    ' ��� Text12 ��ֵ�Ƿ�Ϊ��
    If text12Value = "" Then
        ' ���Ϊ�գ�ֱ����ʾ Forms509
        Forms509.Show
    Else
        parts = Split(text12Value, "/")  ' ʹ��б�ָܷ��ַ���
        If UBound(parts) >= 1 Then
            result = parts(1)  ' ��ȡ�ָ������ĵڶ���Ԫ��
        Else
            result = "No valid data after '/'"  ' ��������Ϊ���ʵ�Ĭ��ֵ�������Ϣ
        End If

        ' �������Ĵ��룬���縳ֵ�� Forms509 �Ŀؼ���
        ' ȷ�� Forms509 �� Text1(3) �ؼ�����
        If Not Forms509 Is Nothing Then
            Forms509.Text1(3).Text = result
            Forms509.Show
        Else
            MsgBox "Forms509 is not loaded or does not exist."
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Label17.Caption = Format(Month(Date), "0#")
Label20.Caption = Format(Month(Date) - 1, "0#")
For i = 1 To 35
Label1(i).Caption = ""
Label1(i).Visible = False
Next

If yhxm <> "" Then
Option3.value = True
Option4.Visible = False
Else
Option4.value = True
Option3.Visible = False
End If

Option2.value = True
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
sjsx = 0
If InStr(yhdm, "2") > 0 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select ��������ϵ�� from gyshd where ��������ϵ��<>'0' group by ��������ϵ�� order by ��������ϵ��"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
i = 1
Adodc7.Recordset.MoveFirst
Do While Not Adodc7.Recordset.EOF
Label13(i).Caption = Adodc7.Recordset.Fields(0)
Label13(i).Visible = True
i = i + 1
Adodc7.Recordset.MoveNext
Loop
End If


Label11.Caption = "��ɨ�����̿�"
Text2.TabIndex = 0
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(1) = 1600
VSFlexGrid1.ColWidth(2) = 1600

ActivateKeyboardLayout 134481924, 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(�û�,�˵�,���) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
sql2 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
Text8.Text = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "-") + 1, InStr(Label1(Index).Caption, "/") - InStr(Label1(Index).Caption, "-") - 1)
Text3.Text = Mid(Label1(Index).Caption, 1, InStr(Label1(Index).Caption, "-") - 1)
l1 = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "-") + 1, InStr(Label1(Index).Caption, "/") - InStr(Label1(Index).Caption, "-") - 1)
L2 = Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, "/") + 1)
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select ���� from v_lv_sczd where ����='" & l1 & "' order by ��ʼ"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF And Adodc11.Recordset.Fields(0) = Text1 Then
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), Text3.Text) > 0 Then
List1.Selected(i) = True
End If
Next
End If
If InStr(yhdm, "2") > 0 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
End Select
End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label1(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label1(Index).BackColor = &HC0FFC0
End Select
End Sub

Private Sub Label10_Click()
'If Val(Text3.Text) > 1000 And Val(Text3.Text) < 6000 Then Exit Sub
Text7.Text = ""
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.BackColor = &H8080FF
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.BackColor = &HC0FFC0
End Sub

Private Sub Label13_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
Text9.Text = ""
Text8.Text = ""
Text3.Text = ""
Label3.Caption = Label13(Index).Caption
Text10.Text = Label13(Index).Caption
For i = 1 To 36
Label1(i).Caption = ""
Label1(i).Visible = False
Next
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select * from gyshd where ��������ϵ��='" & Label13(Index).Caption & "'  order by ���ձ��"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
i = 1
Adodc7.Recordset.MoveFirst
Do While Not Adodc7.Recordset.EOF
Label1(i).Caption = Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + Adodc7.Recordset.Fields(2)
Label1(i).Visible = True
i = i + 1
Adodc7.Recordset.MoveNext
Loop
End If
End Select
End Sub

Private Sub Label13_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label13(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label13_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label13(Index).BackColor = &HFFFF00
End Select
End Sub

Private Sub Label15_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Option1.value = True Then
Text11.Text = Text11.Text + Label15(Index).Caption
End If
If Option2.value = True Then
Text2.Text = Text2.Text + Label15(Index).Caption
End If
End Select
End Sub

Private Sub Label15_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label15(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label15_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label15(Index).BackColor = &HFFFFC0
End Select
End Sub

Private Sub Label16_Click()
On Error Resume Next
If Option1.value = True Then
Text11.Text = Mid(Text11, 1, Len(Text11) - 1)
End If
If Option2.value = True Then
Text2.Text = Mid(Text2, 1, Len(Text2) - 1)
End If
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &H8080FF
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &HFFFFC0
End Sub

Private Sub Label17_Click()
Text2 = "A" + Mid(Format(Date, "YYYY"), 3) + Label17.Caption
End Sub






Private Sub Label18_Click()
pmbl = 7
Formr440.Show
End Sub

Private Sub Label19_Click()
Forms512.Show
End Sub

Private Sub Label20_Click()
Text2 = "A" + Mid(Format(Date, "YYYY"), 3) + Label20.Caption
End Sub

Private Sub Label23_Click()
Text2 = Text12 + "J"
End Sub

Private Sub Label3_Click()
YGBL = 9
Forms546.Text1(0) = Label3.Caption
Forms546.Show
End Sub

Private Sub Label9_Click(Index As Integer)
Select Case Index
       Case Index
Text7.Text = Text7.Text + Label9(Index).Caption
End Select
End Sub


Private Sub Label9_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label9(Index).BackColor = &H8080FF
End Select
End Sub

Private Sub Label9_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
       Case Index
Label9(Index).BackColor = &HC0FFC0
End Select
End Sub

Private Sub Text1_Change()
On Error Resume Next ' ��������������󣬼���ִ�У�����ʾ������Ϣ

' ���� Adodc8 �����ݿ������ַ���
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

' ���� Adodc1 �����ݿ������ַ�������ѯ kpd ������ Text1.Text ƥ��ļ�¼
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ����,Ʒ��,ɫ��,ƥ��,����,��̨,ztbh as ������� from kpd where ����='" & Text1.Text & "' "
Adodc1.Refresh ' ˢ������Դ

' ���� Adodc5 �����ݿ������ַ�������ѯ ghgx ��ɸѡ�ض�����Χ�ļ�¼
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select distinct ���� from ghgx where ����='" & Text1.Text & "' and ���� not between '0080' and '0200'  order by ����"
Adodc5.Refresh ' ˢ������Դ

sjsx = 0 ' ��ʼ������ sjsx
List1.Clear ' ��� List1 �б��
For i = 1 To 35 ' ���� 1 �� 35 �ű�ǩ
    Label1(i).Caption = "" ' ��ձ�ǩ����ʾ����
    Label1(i).Visible = False ' ���ر�ǩ
Next

If Adodc5.Recordset.EOF Then ' ��� Adodc5 �����Ϊ��
    For i = 1 To 35 ' �ٴ�������б�ǩ
        Label1(i).Caption = ""
        Label1(i).Visible = False
    Next
Else ' ��� Adodc5 �������������
    Adodc5.Recordset.MoveFirst ' �ƶ�����һ����¼
    i = 1 ' ��ʼ�������� i
    Do While Not Adodc5.Recordset.EOF ' �������м�¼
        ' ���� Adodc7 �����ݿ������ַ�������ѯ gyshd �����빤��ƥ��ļ�¼
        Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
        Adodc7.RecordSource = "select * from gyshd where ���ձ��='" & Adodc5.Recordset.Fields(0) & "'"
        Adodc7.Refresh ' ˢ������Դ

        If Not Adodc7.Recordset.EOF Then ' ��� Adodc7 �������������
            If InStr(yhdm, "1") > 0 Then ' ��� yhdm �ַ����а��� '1'
                ' ��鹤�ձ���Ƿ����ض���Χ��
                If Adodc7.Recordset.Fields(0) > 0 And Adodc7.Recordset.Fields(0) < 900 Then
                    ' ���ñ�ǩ����ʾ����Ϊ���ձ�š����Ƽ��� '1'
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True ' ��ʾ��ǩ

                    ' ��ѯ pbcl ��������ź͹��ձ��ƥ��ļ�¼
                    Adodc8.RecordSource = "select * from pbcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh ' ˢ������Դ

                    If Adodc8.Recordset.EOF Then ' ��� pbcl ����û��ƥ���¼
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1" ' ����¼��ӵ��б��
                        Label1(i).Enabled = True ' ���ñ�ǩ
                    Else
                        ' ��ѯ kpd ��������ź͹���ƥ�������
                        Adodc6.RecordSource = "select round(sum(����),2) from kpd where ����='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh ' ˢ������Դ

                        ' ��ѯ pbcl ��������ź͹��ձ��ƥ��İ�β���
                        Adodc8.RecordSource = "select round(sum(��β���),2) from pbcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh ' ˢ������Դ

                        ' ��� kpd ���е�����С�ڻ���� pbcl ���еİ�β��������ñ�ǩ
                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else ' �������ñ�ǩ������¼��ӵ��б��
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1 ' ���Ӽ�����
                End If
            End If

            ' ��� yhdm ���Ƿ���� "2"
            If InStr(yhdm, "2") > 0 Then
                ' ��鹤�ձ���Ƿ����ض���Χ��
                If Adodc7.Recordset.Fields(0) > 1000 And Adodc7.Recordset.Fields(0) < 6000 Then
                    ' ���ñ�ǩ���ݲ���ѯ rscl ��
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from rscl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        ' ��ѯ pld ��������� rscl ��İ�β���
                        Adodc6.RecordSource = "select round(sum(����),2) from pld where ����='" & Text1.Text & "' and ��Ϣ like '%����%'"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(��β���),2) from rscl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        ' �Ա� pld �� rscl ���е������Ͳ���
                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If

            ' ��� yhdm ���Ƿ���� "3" (����������ˮ����)
            If InStr(yhdm, "3") > 0 Then
                If Adodc7.Recordset.Fields(0) > 6001 And Adodc7.Recordset.Fields(0) < 7000 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from tscl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(����),2) from kpd where ����='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(��β���),2) from tscl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If

            ' ��� yhdm ���Ƿ���� "4" (���ں�ɹ���)
            If InStr(yhdm, "4") > 0 Then
                If Adodc7.Recordset.Fields(0) > 7000 And Adodc7.Recordset.Fields(0) < 8000 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from hgcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(����),2) from kpd where ����='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(��β���),2) from hgcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
           
            
            If Adodc7.Recordset.Fields(0) > 8000 And Adodc7.Recordset.Fields(0) < 9000 Then
         Label1(i).Caption = Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
         Label1(i).Visible = True

         Adodc8.RecordSource = "select * from xdcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
         Adodc8.Refresh
         If Adodc8.Recordset.EOF Then
        List1.AddItem Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
        Label1(i).Enabled = True

         Else
        Adodc6.RecordSource = "select round(sum(����),2) from kpd where ����='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
         Adodc6.Refresh
       Adodc8.RecordSource = "select round(sum(��β���),2) from xdcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
        Adodc8.Refresh

       If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
        Label1(i).Enabled = False
    Else
       List1.AddItem Adodc7.Recordset.Fields(0) + "-" + Adodc7.Recordset.Fields(1) + "/" + "1"
        Label1(i).Enabled = True

        End If
       End If
       i = i + 1
      End If
      End If

            ' ��� yhdm ���Ƿ���� "5" (���ڴ��͹���)
            If InStr(yhdm, "5") > 0 Then
                If Adodc7.Recordset.Fields(0) > 9000 And Adodc7.Recordset.Fields(0) < 9999 Then
                    Label1(i).Caption = Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                    Label1(i).Visible = True
                    Adodc8.RecordSource = "select * from ddcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                    Adodc8.Refresh

                    If Adodc8.Recordset.EOF Then
                        List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                        Label1(i).Enabled = True
                    Else
                        Adodc6.RecordSource = "select round(sum(����),2) from kpd where ����='" & Text1.Text & "' and PATINDEX('%'+'" & Adodc7.Recordset.Fields(0) & "'+'%',GX)>0"
                        Adodc6.Refresh
                        Adodc8.RecordSource = "select round(sum(��β���),2) from ddcl where ����='" & Text1.Text & "' and ���ձ��='" & Adodc7.Recordset.Fields(0) & "'"
                        Adodc8.Refresh

                        If Val(Adodc6.Recordset.Fields(0)) <= Val(Adodc8.Recordset.Fields(0)) Then
                            Label1(i).Enabled = False
                        Else
                            List1.AddItem Adodc7.Recordset.Fields(0) & "-" & Adodc7.Recordset.Fields(1) & "/" & "1"
                            Label1(i).Enabled = True
                        End If
                    End If
                    i = i + 1
                End If
            End If
        End If

        Adodc5.Recordset.MoveNext ' �ƶ�����һ����¼
    Loop
End If

' ����ǩ�Ƿ����ò������ձ�ŵ��Զ�ɨ���״̬����
For i = 1 To 35
    If Label1(i).Enabled = True Then
        L = i
        dqgx = L ' ��ǰ����Ϊ L

        If Option3.value = True Then
            If Mid(List1.List(0), 1, 4) = yhxm And Text12 <> "" Then
                ' ������ձ�ŵ����û����������Ұ����Ա����Ϣ��Ϊ�գ����Զ�ɨ�����
                Label1_Click (dqgx)
                Text2 = Text12 & "J"
            End If
        End If

        GoTo 100
    End If
Next

100:

' �����ǩ�ĵ�һ���ַ��� "12345"����������Ӧ�ĺ�����ǩ
If InStr("12345", Left(Label1(L).Caption, 1)) > 0 Then
    For m = L + 1 To 35
        If InStr("12345", Left(Label1(m).Caption, 1)) > 0 Then
            Label1(m).Enabled = True
        Else
            Label1(m).Enabled = True '''''����ĳ�flaseɨ����һ���������ɨ����Ĺ���
        End If
    Next
Else
    ' ���򣬽��ú�����ǩ
    For m = L + 1 To 35
        Label1(m).Enabled = True '''''����ĳ�flaseɨ����һ���������ɨ����Ĺ���
    Next
End If
End Sub



Private Sub Text2_Change()
On Error Resume Next

If InStr(Text2.Text, "J") > 0 Or InStr(Text2.Text, "j") > 0 Then

m = Mid(Text2.Text, 1, Len(Text2.Text) - 1)

'If Len(M) = 7 And Mid(M, 1, 1) = "8" Then
If (InStr(m, ".") > 0 And Len(m) / 4 = Int(Len(m) / 4)) Or (InStr(m, ".") > 0 And InStr(m, "/") > 0) Then
If Adodc1.Recordset.EOF Then
Label11.Caption = "��ɨ�����̿�"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text8.Text = ""
Text2.SetFocus
Else

If InStr(m, "/") > 0 Then

bz = Mid(m, 1, InStr(m, "/") - 1)       '����
'BZ = yhdm      '����
yg = Mid(m, InStr(m, "/") + 1)          'Ա��
'yg = M          'Ա��

Else
'bz = Mid(M, 1, InStr(M, "/") - 1)       '����
bz = bzdm       '����
'yg = Mid(M, InStr(M, "/") + 1)          'Ա��
yg = m          'Ա��
End If

If Text3.Text <> "" And Len(Text3.Text) = 4 Then

If Val(Text7.Text) <= 0 Then
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Exit Sub
End If

djsj = Now

For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then

Text8.Text = Mid(List1.List(i), InStr(List1.List(i), "-") + 1, InStr(List1.List(i), "/") - InStr(List1.List(i), "-") - 1)
Text3.Text = Mid(List1.List(i), 1, InStr(List1.List(i), "-") - 1)    ''''''������


Adodc10.RecordSource = "select ���� from ghgx where ����='" & Text1 & "' and ����>'" & Text3 & "' order by ����"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
If Val(Adodc10.Recordset.Fields(0)) < 1000 Or Val(Adodc10.Recordset.Fields(0)) > 6000 Then
gybb = Adodc10.Recordset.Fields(0)
sql1 = "update ghgx  set ��ʼ='" & Now & "' where ����='" & Text1 & "' and ����='" & Adodc10.Recordset.Fields(0) & "'"
sql2 = "update ghgx  set ����='" & Now & "' where ����='" & Text1 & "' and ����='" & Text3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Else
gybb = "9999"
sql1 = "update ghgx  set ����='" & Now & "' where ����='" & Text1 & "' and ����='" & Text3 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

l1 = Mid(List1.List(i), InStr(List1.List(i), "-") + 1, InStr(List1.List(i), "/") - InStr(List1.List(i), "-") - 1)
L2 = Mid(List1.List(i), InStr(List1.List(i), "/") + 1)
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select ��������ϵ��,������ϵ��,�������� from gyshd where ���ձ��='" & Text3 & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
l1 = Adodc7.Recordset.Fields(2)   ''''��������
L2 = Adodc7.Recordset.Fields(1)   ''''��������ϵ��
L3 = Adodc7.Recordset.Fields(0)   ''''������ϵ��
Else
l1 = l1
L2 = 0
L3 = ""
End If

If Text8 <> l1 Then
MsgBox ("��ȷ�Ϲ���ѡ�� �Ƿ���ȷ")
Exit Sub
End If

'If Val(Text3) < 1000 And Val(Adodc10.Recordset.Fields(0)) > 1000 And Val(Adodc10.Recordset.Fields(0)) < 6000 Then
'gybb = Adodc10.Recordset.Fields(0)
'l1 = "Ⱦ�״���"
'End If

If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then       ''''''''''''''''''�Ų�

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmpb1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"    ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute  ' ִ�д洢����        ����               ������          ��������      ����         Ա��          ʱ��       ��������ϵ��     δ����                �����         ������ϵ��       ��̨            ������
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then    '''''''''''''''''Ⱦɫ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ӱ���
Adodc9.RecordSource = "select distinct ���� from ghgx where ����='" & Text1 & "' and ����='" & Text3 & "'"
Adodc9.Refresh
If Not Adodc9.Recordset.EOF Then
bs = Val(Adodc9.Recordset.Fields(0))
End If

If InStr(l1, "����") > 0 Then
Adodc10.RecordSource = "select ���� from ghgx where ����='" & Text1 & "' and ����>'6000' order by ����"
Adodc10.Refresh
If Not Adodc10.Recordset.EOF Then
gybb = Adodc10.Recordset.Fields(0)
Else
gybb = Text3
End If
End If

If Val(L2) > 0.2 Then '''''�������ϵ������2ëǮ
L2 = L2 * bs                '''' װ������

If hgxx = "1" Then                 '''''''''''''''''''''û�кϸ׵�ִ��cjsmrsa
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmrsa('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus

Else    ''''''�кϸ׵�ִ��cjsmrsc,���ɨ��Ĺ��������ˣ�����Ϊ�ϸ���,�ְѺϸ׵��ϵ�ɾ����

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmrsc('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & hgxx & "','" & L3 & "','" & Text11 & "','" & gybb & "')"    ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If



Else '''''����ϵ��С��2ëǮ

L2 = L2 * bs                '''' װ������
If hgxx = "1" Then                 '''''''''''''''''''''�ϸ���Ϣ
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmrsb('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmrsd('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & hgxx & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

End If
End If


If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then    '''''''''''''''''��ˮ
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmts1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then    '''''''''''''''''���
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmhg1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then    '''''''''''''''''С����
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmxd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then    '''''''''''''''''����
If InStr(Text8.Text, "��������") > 0 Then
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select ���� from jgmx where ����='" & Text1 & "'"
Adodc2.Refresh

If Adodc2.Recordset.EOF Then
MsgBox ("û�п��߷������ݣ����ܳ���")
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
Exit Sub
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmdd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If
Else
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "cjsmdd1('" & Text1.Text & "','" & Text3.Text & "','" & l1 & "','" & bz & "','" & yg & "','" & djsj & "','" & L2 & "','" & Text6.Text & "','" & Text7.Text & "','" & L3 & "','" & Text11 & "','" & gybb & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
Adodc3.Refresh
'Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
'Text8.Text = ""
Text2.SetFocus
End If
End If

End If     'listѡȡ��
Next
Call Text1_Change

Else
Text2.Text = ""
Text2.SetFocus
Exit Sub
End If


Label11.Caption = "��ɨ�����̿�"
Text2.Text = ""
Exit Sub
End If
End If

If Len(m) > 3 Then
Text2.Text = ""
Text3.Text = ""
' ���m���Ƿ����"+", ������ڣ������滻Ϊ���ַ�����ȥ�����е�"+"
If InStr(m, "+") > 0 Then
    Text1.Text = Replace(m, "+", "")
Else
    Text1.Text = m
End If
Label11.Caption = "��ѡ����"
End If

End If

End Sub

Private Sub Text3_Change()
On Error Resume Next
If Val(Text3.Text) >= 1 And Val(Text3.Text) <= 1000 Then
Adodc3.RecordSource = "select * from PBCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 1)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "smccjc1"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value) '''
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)

End If

' ���Text3�ؼ����ı�ֵ���ڵ���1001��С�ڵ���6000����ִ�����²���
If Val(Text3.Text) >= 1001 And Val(Text3.Text) <= 6000 Then

    ' ����Adodc4�ؼ�������ԴΪһ��SQL��ѯ��䣬��ѯbgxx�����������ϱ�ţ������ǲ��׹��ŵ���Text1�ؼ����ı�ֵ
    Adodc4.RecordSource = "select ���ϱ�� from bgxx where ���׹���='" & Text1.Text & "'"
    ' ˢ��Adodc4�ؼ���ִ�в�ѯ
    Adodc4.Refresh
    ' ���Adodc4�ļ�¼��������ĩβ��˵��û���ҵ���¼
    If Adodc4.Recordset.EOF Then
        ' ����hgxx������ֵΪ"1"
        hgxx = "1"
        ' ����Adodc3�ؼ�������ԴΪһ��SQL��ѯ��䣬��ѯRSCL��������У������ǹ��ŵ���Text1�ؼ����ı�ֵ�����ձ�ŵ���Text3�ؼ����ı�ֵ
        Adodc3.RecordSource = "select * from RSCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
        ' ˢ��Adodc3�ؼ���ִ�в�ѯ
        Adodc3.Refresh
    Else
        ' ���Adodc4�ļ�¼��û�е���ĩβ��˵���ҵ��˼�¼
        ' ��ȡAdodc4��¼���ĵ�һ���ֶ�ֵ������hgxx����
        hgxx = Adodc4.Recordset.Fields(0)
        ' ��������Adodc3�ؼ�������Դ����������ͬ�Ĳ�ѯ
        Adodc3.RecordSource = "select * from RSCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
        ' ˢ��Adodc3�ؼ���ִ�в�ѯ
        Adodc3.Refresh
    End If

    ' ����һ���µ�Command��������ִ�����ݿ�����
    Set g_Cmd = New Command
    ' �������ݿ������ַ������������ݿ��ṩ�ߡ����롢�û�ID�����ݿ���������Դ��ַ
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    ' ��Command����Ļ��������Ϊ����������ַ���
    g_Cmd.ActiveConnection = g_Con

    ' �����Ǵ�����׷�Ӷ��������Command�����Թ��洢����ʹ��
    ' ����һ����������ʾ���ţ��������ͣ�����40��ֵΪText1���ı�ֵȥ�����˿ո�
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text)) '''�׺�
    g_Cmd.Parameters.Append param

    ' ����һ����������ʾ���ձ�ţ��������ͣ�����4��ֵΪText3���ı�ֵȥ�����˿ո�
    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text)) ''''���ձ��
    g_Cmd.Parameters.Append param
    
    ' ����Adodc4�Ƿ��ҵ���¼�����ò�ͬ������ֵ����tj������
    ' ���Adodc4û���ҵ���¼��"tj"������ֵΪ2������Ϊ8
    If Adodc4.Recordset.EOF Then
        Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 2)
    Else
        Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 8)
    End If
    g_Cmd.Parameters.Append param

    ' ������������������ֱ���"pb"��"cl"��"pb1"�����ǵ����Ͷ��ǵ����ȸ�����
    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    ' ����Command���������Ϊ�洢���̣���ָ���洢���̵�����Ϊ"smccjc"
    g_Cmd.CommandType = adCmdStoredProc
    g_Cmd.CommandText = "smccjc"
    ' ִ�д洢����
    g_Cmd.Execute
    ' ȡ��Command�����ִ�У���������
    g_Cmd.Cancel

    ' ��Command����Ĳ��������л�ȡ���������ֵ����ʾ����Ӧ���ı�����
    Text4.Text = Val(g_Cmd.Parameters("pb").value)  ' "pb"������ֵ��ʾ��Text4��
    Text5.Text = Val(g_Cmd.Parameters("cl").value)  ' "cl"������ֵ��ʾ��Text5��
    ' ���㲢��ʾText4��Text5�Ĳ�ֵ��Text6��Text7��
    Text6.Text = Val(Text4.Text) - Val(Text5.Text)
    Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If


If Val(Text3.Text) >= 6001 And Val(Text3.Text) <= 7000 Then
Adodc3.RecordSource = "select * from TSCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 3)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "smccjc1"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value)
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If


If Val(Text3.Text) >= 7001 And Val(Text3.Text) <= 8000 Then
Adodc3.RecordSource = "select * from HGCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 4)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "smccjc1"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel



Text4.Text = Val(g_Cmd.Parameters("pb").value) ''�Ų���
Text5.Text = Val(g_Cmd.Parameters("cl").value) ''�ѱ���
Text6.Text = Val(Text4.Text) - Val(Text5.Text) '' δ����
Text7.Text = Val(Text4.Text) - Val(Text5.Text) ''�����

End If

If Val(Text3.Text) >= 8001 And Val(Text3.Text) <= 9000 Then
Adodc3.RecordSource = "select * from XDCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 5)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    
    Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
  
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "smccjc1"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value) ''�Ų���
Text5.Text = Val(g_Cmd.Parameters("cl").value) ''�ѱ���
Text6.Text = Val(Text4.Text) - Val(Text5.Text) '' δ����
Text7.Text = Val(Text4.Text) - Val(Text5.Text) ''�����
End If

If Val(Text3.Text) >= 9001 And Val(Text3.Text) <= 9999 Then
Adodc3.RecordSource = "select * from DDCL where ����='" & Text1.Text & "' and ���ձ��='" & Text3.Text & "'"
Adodc3.Refresh
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(Text1.Text))
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("gy", adChar, adParamInput, 4, Trim(Text3.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("gx", adChar, adParamInput, 20, Trim(Text8.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("tj", adChar, adParamInput, 2, 6)
    g_Cmd.Parameters.Append param

    Set param = g_Cmd.CreateParameter("pb", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("cl", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
     Set param = g_Cmd.CreateParameter("pb1", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
   
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "smccjc1"       ' ��ʾ�����ĸ��洢����"
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Text4.Text = Val(g_Cmd.Parameters("pb").value)
Text5.Text = Val(g_Cmd.Parameters("cl").value)
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text7.Text = Val(Text4.Text) - Val(Text5.Text)
End If

VSFlexGrid3.ColWidth(0) = 100
VSFlexGrid3.ColWidth(1) = 0
VSFlexGrid3.ColWidth(2) = 0
VSFlexGrid3.ColWidth(6) = 0
VSFlexGrid3.ColWidth(10) = 0
VSFlexGrid3.ColWidth(12) = 0
VSFlexGrid3.ColWidth(13) = 0
VSFlexGrid3.ColWidth(14) = 0
VSFlexGrid3.ColWidth(15) = 0
VSFlexGrid3.ColWidth(16) = 0
VSFlexGrid3.ColWidth(17) = 0
VSFlexGrid3.ColWidth(18) = 0

VSFlexGrid1.ColWidth(0) = 200
Label11.Caption = "��ɨ�蹤������"

End Sub

Private Sub Timer1_Timer()
If sjsx >= 1 Then
Text12 = bzgrbh
Timer1.Enabled = False
sjsx = 1
Else
sjsx = sjsx + 1
End If
End Sub
