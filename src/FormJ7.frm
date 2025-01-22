VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formj7 
   BackColor       =   &H00C0E0FF&
   Caption         =   "染机排缸"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   12855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   23535
      _ExtentX        =   41513
      _ExtentY        =   22675
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      BackColor       =   12180727
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "排产信息"
      TabPicture(0)   =   "FormJ7.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "待排信息"
      TabPicture(1)   =   "FormJ7.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   12015
         Left            =   -75000
         ScaleHeight     =   11955
         ScaleWidth      =   23355
         TabIndex        =   52
         Top             =   720
         Width           =   23415
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "FormJ7.frx":0038
            Left            =   9960
            List            =   "FormJ7.frx":0042
            TabIndex        =   113
            Text            =   "Combo2"
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            Height          =   375
            Index           =   2
            Left            =   3600
            TabIndex        =   106
            Text            =   "Text16"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text16 
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   105
            Text            =   "Text16"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text16 
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   104
            Text            =   "Text16"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text15 
            Height          =   375
            Index           =   2
            Left            =   3600
            TabIndex        =   103
            Text            =   "Text15"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox Text15 
            Height          =   375
            Index           =   1
            Left            =   3000
            TabIndex        =   102
            Text            =   "Text15"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox Text15 
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   101
            Text            =   "Text15"
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "状态"
            Height          =   375
            Index           =   13
            Left            =   13560
            TabIndex        =   99
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "锅号"
            Height          =   375
            Index           =   11
            Left            =   12840
            TabIndex        =   81
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "客户"
            Height          =   375
            Index           =   10
            Left            =   12840
            TabIndex        =   80
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   8040
            TabIndex        =   79
            Text            =   "Text11"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            Height          =   375
            Left            =   5520
            TabIndex        =   76
            Text            =   "Text7"
            Top             =   960
            Width           =   615
         End
         Begin VB.Frame framel1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "排产修改栏"
            Height          =   975
            Left            =   15480
            TabIndex        =   69
            Top             =   360
            Visible         =   0   'False
            Width           =   3255
            Begin VB.TextBox Text9 
               Height          =   270
               Left            =   720
               TabIndex        =   73
               Text            =   "Text9"
               Top             =   600
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox Text8 
               Height          =   270
               Left            =   720
               TabIndex        =   71
               Text            =   "Text8"
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label21 
               BackColor       =   &H00C0C0FF&
               Caption         =   "取消"
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
               Left            =   2640
               TabIndex        =   83
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label18 
               BackColor       =   &H00C0C0FF&
               Caption         =   "确认"
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
               Left            =   2640
               TabIndex        =   74
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label17 
               BackColor       =   &H0000C0C0&
               Caption         =   "色别"
               Height          =   255
               Left            =   240
               TabIndex        =   72
               Top             =   600
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label16 
               BackColor       =   &H0000C0C0&
               Caption         =   "锅号"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "日期"
            Height          =   375
            Index           =   9
            Left            =   12120
            TabIndex        =   66
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "重量"
            Height          =   375
            Index           =   8
            Left            =   12120
            TabIndex        =   59
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   6720
            TabIndex        =   58
            Text            =   "Text4"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   5520
            TabIndex        =   57
            Text            =   "Text5"
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "查询"
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
            Left            =   14280
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   480
            Width           =   1215
         End
         Begin VB.Data Data2 
            Caption         =   "Data2"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  '缺省游标
            DefaultType     =   2  '使用 ODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   600
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   10320
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  '缺省游标
            DefaultType     =   2  '使用 ODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   600
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   10080
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "打印"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14280
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   960
            Width           =   1215
         End
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   375
            Left            =   6000
            Top             =   10440
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
         Begin MSAdodcLib.Adodc Adodc8 
            Height          =   375
            Left            =   6360
            Top             =   9840
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
            Bindings        =   "FormJ7.frx":0052
            Height          =   9855
            Left            =   600
            TabIndex        =   55
            Top             =   1560
            Width           =   18375
            _cx             =   32411
            _cy             =   17383
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
            FormatString    =   $"FormJ7.frx":0067
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
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1080
            TabIndex        =   62
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarTitleBackColor=   8421440
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   329580547
            CurrentDate     =   39961
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   375
            Left            =   1080
            TabIndex        =   63
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarTitleBackColor=   8421440
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   329580547
            CurrentDate     =   39961
         End
         Begin MSDataListLib.DataCombo DataCombo9 
            Bindings        =   "FormJ7.frx":02BE
            Height          =   330
            Index           =   0
            Left            =   6120
            TabIndex        =   77
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "简称"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label19 
            Caption         =   "状态"
            Height          =   375
            Index           =   2
            Left            =   9960
            TabIndex        =   100
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "锅号"
            Height          =   375
            Index           =   1
            Left            =   8040
            TabIndex        =   78
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "客户"
            Height          =   375
            Index           =   0
            Left            =   4560
            TabIndex        =   75
            Top             =   960
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   6480
            X2              =   6720
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Index           =   12
            Left            =   600
            TabIndex        =   65
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Index           =   11
            Left            =   600
            TabIndex        =   64
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "重量范围"
            Height          =   255
            Left            =   4560
            TabIndex        =   60
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   10215
         Left            =   0
         ScaleHeight     =   10155
         ScaleWidth      =   18915
         TabIndex        =   1
         Top             =   720
         Width           =   18975
         Begin VB.TextBox Text18 
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   112
            Text            =   "Text18"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   111
            Text            =   "Text18"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text18 
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   110
            Text            =   "Text18"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   109
            Text            =   "Text17"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   108
            Text            =   "Text17"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   107
            Text            =   "Text17"
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "FormJ7.frx":02D3
            Left            =   7920
            List            =   "FormJ7.frx":02E0
            TabIndex        =   98
            Text            =   "Combo1"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "排缸方式"
            Height          =   615
            Left            =   7440
            TabIndex        =   91
            Top             =   7200
            Width           =   3735
            Begin VB.OptionButton Option4 
               Caption         =   "排完"
               Height          =   255
               Left            =   2160
               TabIndex        =   93
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option3 
               Caption         =   "续排"
               Height          =   255
               Left            =   720
               TabIndex        =   92
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.TextBox Text14 
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
            Left            =   12720
            TabIndex        =   90
            Text            =   "Text14"
            Top             =   6720
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "编号方式"
            Height          =   495
            Left            =   15480
            TabIndex        =   86
            Top             =   5160
            Width           =   3015
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C0FFC0&
               Caption         =   "自动"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   7.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1080
               TabIndex        =   88
               Top             =   120
               Width           =   615
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C0FFC0&
               Caption         =   "手动"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   7.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2160
               TabIndex        =   87
               Top             =   120
               Width           =   615
            End
         End
         Begin VB.TextBox Text13 
            Height          =   375
            Left            =   16200
            TabIndex        =   85
            Text            =   "Text13"
            Top             =   6000
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Left            =   8400
            TabIndex        =   84
            Text            =   "Text12"
            Top             =   6720
            Width           =   495
         End
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   13800
            Top             =   9240
         End
         Begin MSAdodcLib.Adodc Adodc11 
            Height          =   330
            Left            =   11040
            Top             =   9360
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            Height          =   375
            Left            =   9000
            Top             =   9360
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
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   5280
            TabIndex        =   67
            Text            =   "Text7"
            Top             =   240
            Width           =   615
         End
         Begin MSAdodcLib.Adodc Adodc9 
            Height          =   330
            Left            =   5880
            Top             =   9480
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
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "打印"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   17040
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "查询"
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
            Left            =   17040
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "退出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   17040
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "查询条件"
            Height          =   1095
            Left            =   12240
            TabIndex        =   6
            Top             =   120
            Width           =   4935
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "班次"
               Height          =   255
               Index           =   12
               Left            =   3840
               TabIndex        =   97
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "未配料"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   14
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "锅号"
               Height          =   255
               Index           =   6
               Left            =   1800
               TabIndex        =   13
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "机台"
               Height          =   255
               Index           =   7
               Left            =   960
               TabIndex        =   12
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "操作"
               Height          =   255
               Index           =   5
               Left            =   1800
               TabIndex        =   11
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "日期"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "颜色"
               Height          =   255
               Index           =   3
               Left            =   960
               TabIndex        =   9
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "染色中"
               Height          =   255
               Index           =   2
               Left            =   2640
               TabIndex        =   8
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "客户"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   7
               Top             =   720
               Width           =   735
            End
         End
         Begin VB.TextBox Text1 
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
            Left            =   9000
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   6000
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   9000
            TabIndex        =   4
            Text            =   "Text2"
            Top             =   8280
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   0
            Top             =   3000
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   12720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "FormJ7.frx":02F0
            Top             =   6000
            Width           =   2175
         End
         Begin VB.TextBox Text4 
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
            Left            =   16200
            TabIndex        =   2
            Text            =   "Text2"
            Top             =   6720
            Width           =   2175
         End
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   330
            Left            =   7080
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   375
            Left            =   7560
            Top             =   10440
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
            Left            =   7920
            Top             =   10560
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   330
            Left            =   7920
            TabIndex        =   18
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo4"
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   330
            Left            =   5640
            TabIndex        =   19
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo2"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FormJ7.frx":02F6
            Height          =   330
            Left            =   5880
            TabIndex        =   20
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "简称"
            Text            =   "DataCombo1"
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1200
            TabIndex        =   21
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarTitleBackColor=   8421440
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   329646083
            CurrentDate     =   39961
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1200
            TabIndex        =   22
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarTitleBackColor=   8421440
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   329646083
            CurrentDate     =   39961
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "FormJ7.frx":030B
            Height          =   330
            Left            =   10200
            TabIndex        =   23
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "姓名"
            Text            =   "DataCombo4"
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "FormJ7.frx":0321
            Height          =   330
            Left            =   10200
            TabIndex        =   24
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "车台编号"
            Text            =   "DataCombo4"
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   375
            Left            =   7800
            Top             =   10440
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
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "FormJ7.frx":0336
            Height          =   390
            Left            =   9000
            TabIndex        =   25
            Top             =   6720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   688
            _Version        =   393216
            Style           =   2
            ListField       =   "车台编号"
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   375
            Left            =   7560
            Top             =   10440
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
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   12720
            TabIndex        =   26
            Top             =   5280
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   329646083
            CurrentDate     =   36892
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "FormJ7.frx":034B
            Height          =   390
            Left            =   16680
            TabIndex        =   27
            Top             =   6000
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "车台编号"
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   375
            Left            =   7920
            Top             =   10560
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Bindings        =   "FormJ7.frx":0360
            Height          =   3375
            Left            =   360
            TabIndex        =   28
            Top             =   5280
            Width           =   6735
            _cx             =   11880
            _cy             =   5953
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
            FormatString    =   $"FormJ7.frx":0375
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
            Bindings        =   "FormJ7.frx":044E
            Height          =   615
            Left            =   360
            TabIndex        =   56
            Top             =   8640
            Width           =   6735
            _cx             =   11880
            _cy             =   1085
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
            FormatString    =   $"FormJ7.frx":0463
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
            Bindings        =   "FormJ7.frx":053C
            Height          =   3615
            Left            =   360
            TabIndex        =   68
            Top             =   1320
            Width           =   18135
            _cx             =   31988
            _cy             =   6376
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
            FormatString    =   $"FormJ7.frx":0551
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
         Begin MSDataListLib.DataCombo DataCombo10 
            Bindings        =   "FormJ7.frx":07AA
            Height          =   390
            Left            =   12720
            TabIndex        =   94
            Top             =   6000
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "名称"
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DataCombo11 
            Bindings        =   "FormJ7.frx":07C0
            Height          =   390
            Left            =   12720
            TabIndex        =   95
            Top             =   6720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   688
            _Version        =   393216
            ListField       =   "姓名"
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0FFC0&
            Caption         =   "配方查询"
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
            Left            =   17400
            TabIndex        =   96
            Top             =   7440
            Width           =   1095
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0FFC0&
            Caption         =   "染缸操作"
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
            Left            =   11520
            TabIndex        =   89
            Top             =   6720
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "刷新"
            Height          =   375
            Left            =   10800
            TabIndex        =   82
            Top             =   6000
            Width           =   375
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFC0&
            Caption         =   "排缸取消"
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
            Left            =   14040
            TabIndex        =   61
            Top             =   7440
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Index           =   6
            Left            =   360
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Index           =   5
            Left            =   360
            TabIndex        =   50
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "客户"
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
            Left            =   4920
            TabIndex        =   49
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "颜色"
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
            Index           =   0
            Left            =   5160
            TabIndex        =   48
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "班次"
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
            Left            =   7440
            TabIndex        =   47
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "锅号"
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
            Left            =   7440
            TabIndex        =   46
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "操作"
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
            Index           =   4
            Left            =   9720
            TabIndex        =   45
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "机台"
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
            Index           =   7
            Left            =   9720
            TabIndex        =   44
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "缸号"
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
            Index           =   9
            Left            =   7440
            TabIndex        =   43
            Top             =   6000
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "排缸机台"
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
            Index           =   10
            Left            =   7440
            TabIndex        =   42
            Top             =   6720
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序设置"
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
            Left            =   14880
            TabIndex        =   41
            Top             =   8280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "排缸确认"
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
            Left            =   12840
            TabIndex        =   40
            Top             =   7440
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
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
            Height          =   375
            Left            =   10680
            TabIndex        =   39
            Top             =   8280
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "车台查询"
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
            Left            =   13680
            TabIndex        =   38
            Top             =   8280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "排产日期"
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
            Index           =   8
            Left            =   11520
            TabIndex        =   37
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "排缸备注"
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
            Left            =   11520
            TabIndex        =   36
            Top             =   6000
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "当前排缸编号"
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
            Left            =   7440
            TabIndex        =   35
            Top             =   5280
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label8"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9000
            TabIndex        =   34
            Top             =   5280
            Width           =   2295
         End
         Begin VB.Label Label9 
            Caption         =   "当前排缸编号"
            Height          =   375
            Left            =   7440
            TabIndex        =   33
            Top             =   8280
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "排缸查询"
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
            Left            =   12480
            TabIndex        =   32
            Top             =   8280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "转缸编号"
            Height          =   375
            Left            =   15480
            TabIndex        =   31
            Top             =   6720
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "转缸机台"
            Height          =   375
            Left            =   15480
            TabIndex        =   30
            Top             =   6000
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0FFC0&
            Caption         =   "转缸确认"
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
            Left            =   16200
            TabIndex        =   29
            Top             =   7440
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "Formj7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fg As VSFlexGrid: Public sj As Integer
Public sql1 As String: Dim cdbhf As Integer
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public jd As Integer

Public S1, S2, R1, R2 As Integer
Private Sub Command2_Click()
Call MXOutadodcToExcel(VSFlexGrid2, "待排计划" + "日期：" + Trim(DTPicker3.value) + "--" + Trim(DTPicker5.value))
End Sub

Private Sub Command5_Click()
sql1 = ""


If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(9).value = 1 Then
t1 = Format(Trim(DTPicker3.value) + Space(2) + Text15(0) + ":" + Text15(1) + ":" + Text15(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker5.value) + Space(2) + Text16(0) + ":" + Text16(1) + ":" + Text16(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "cast(CONVERT(varchar,日期, 120) as datetime) between cast('" & t1 & "' as datetime) and cast('" & t2 & "' as datetime) and "
End If

If Check2(11).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & Text11 & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "色别 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If Check2(8).value = 1 Then
sql1 = sql1 + "合计重量 between '" & Text5 & "' and '" & Text6 & "' and "
End If

If Check2(10).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo9(0) & "'+'%' and "
End If

If Check2(13).value = 1 Then
sql1 = sql1 + "状态 like '%'+'" & Combo2 & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 客户名称,锅号,色别,品名,匹数,重量,类别,染色要求,日期,合计重量,状态,车台  from v_jhkpd where (" + sql1 + ") order by 客户名称, 日期,合计重量"
Adodc8.Refresh

With VSFlexGrid2
    .WordWrap = True
    .MergeCells = 4
    .MergeCol(1) = True '是否上下列合并
    .MergeCol(2) = True '是否上下列合并
    .MergeCol(3) = True '是否上下列合并
    .MergeCol(10) = True '是否上下列合并
End With

VSFlexGrid2.SubtotalPosition = flexSTBelow
VSFlexGrid2.Subtotal flexSTSum, 0, 6, , vbGreen
VSFlexGrid2.Subtotal flexSTCount, 0, 2, , vbGreen
VSFlexGrid2.ColWidth(0) = 200
If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 400
Next
End If
End Sub

Private Sub DataCombo10_Change()
Text3 = DataCombo10
End Sub

Private Sub DataCombo10_Click(Area As Integer)
Text3 = DataCombo10
End Sub

Private Sub DataCombo11_Change()
Text14 = DataCombo11
End Sub

Private Sub DataCombo11_Click(Area As Integer)
Text14 = DataCombo11
End Sub

Private Sub Form_Resize()
On Error Resume Next
  Dim WidthRatio As Double
    Dim HeightRatio As Double
    
    ' Calculate the ratio of the current form size to the original form size
    WidthRatio = Me.Width / 8000
    HeightRatio = Me.Height / 6000
    
    ' Resize and reposition each control based on the current form size
    lblTitle.FontSize = 30 * HeightRatio
    lblTitle.Move 500 * WidthRatio, 500 * HeightRatio, 7000 * WidthRatio, 1000 * HeightRatio
    
    txtName.FontSize = 12 * HeightRatio
    txtName.Move 1000 * WidthRatio, 2000 * HeightRatio, 4000 * WidthRatio, 500 * HeightRatio
    
    txtAddress.FontSize = 12 * HeightRatio
    txtAddress.Move 1000 * WidthRatio, 3000 * HeightRatio, 4000 * WidthRatio, 1000 * HeightRatio
    
    cmdSubmit.FontSize = 12 * HeightRatio
    cmdSubmit.Move 3000 * WidthRatio, 4000 * HeightRatio, 2000 * WidthRatio, 500 * HeightRatio
If Me.WindowState = 1 Then  ''窗体最小化
sql2 = "insert into yhcd(用户,菜单,编号) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"  ''''这行代码执行一个 SQL 语句，将用户、菜单和编号的值插入到一个名为 yhcd 的表中。yhm、Me.Caption 和 cdbhf 是变量或控件的值，它们被连接到 SQL 语句中。注意，这段代码并没有对 SQL 注入进行保护，这可能会导致安全漏洞。
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2  ''将名为 Formm1 的窗体的状态设置为 2（也就是最大化）。
Formm1.Adodc1.Refresh ''刷新名为 Formm1 的窗体上的一个名为 Adodc1 的数据源控件，以显示最新的数据。
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic '''这行代码打开一个名为 RD 的记录集对象，并使用前面的 SQL 语句查询数据库。conn 是一个连接对象，adOpenStatic 和 adLockOptimistic 是记录集的打开类型和锁定类型。
Formm1.Adodc1.Refresh
End Sub

Private Sub Label15_Click()
If MsgBox("确定取消排缸吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE KPD SET ye='N',bh='',bz='',zt='染缸取消',版号='',cky='',班次='',kp1='N',kp='N',RS='N',FH='N',XDX='N' WHERE 锅号='" & Text1.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("排缸取消！")
Call Label5_Click
Call Command5_Click
Call Command1_Click
End Sub

Private Sub Label18_Click()
If Len(Text8) < 3 Then Exit Sub
If MsgBox("确定已排产 锅号" + Text8 + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE KPD SET ye='Y' WHERE 锅号='" & Text8.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("确认成功！")
Call Command5_Click
Call Command1_Click
End Sub

Private Sub Label20_Click()
If Len(Text1) > 3 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_kpdb where  锅号='" & Text1 & "' ORDER BY 车台,排产编号,排产时间"
Adodc1.Refresh
End If
End Sub

Private Sub Label21_Click()
If Len(Text8) < 3 Then Exit Sub
If MsgBox("确定已排产 锅号" + Text8 + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE KPD SET ye='N' WHERE 锅号='" & Text8.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("取消成功！")
Call Command5_Click
Call Command1_Click
End Sub

Private Sub Label23_Click()
Formh224.Show
End Sub

Private Sub Option1_Click()
Text13 = ""
DataCombo8 = ""
Text4 = ""
Text1 = ""
Text3 = ""
Text14 = ""
End Sub

Private Sub Text10_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称  from khzl where 代码 like '%'+'" & Text10 & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub Text12_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 车台编号 from ct where 车台编号 like '%'+'" & Text12 & "'+'%' group by 车台编号"
Adodc3.Refresh
End Sub

Private Sub Text13_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 车台编号 from ct where 车台编号 like '%'+'" & Text13 & "'+'%' group by 车台编号"
Adodc3.Refresh
End Sub

Private Sub Text7_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称  from khzl where 代码 like '%'+'" & Text7 & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub vsfGroup1_GotFocus()

End Sub

Private Sub Timer2_Timer()
If sj = 60 Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "gxsxjc"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
sj = 1
Else
sj = sj + 1
End If
End Sub


Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
R1 = VSFlexGrid1.ColSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
L = 0
For i = S1 To S2
L = L + Val(VSFlexGrid1.TextMatrix(i, R1))
Next
End Sub

Private Sub VSFlexGrid2_Click()
On Error Resume Next
If Adodc8.Recordset.EOF Then Exit Sub
Adodc8.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc8.Recordset.Move rs - 1
Text8 = Adodc8.Recordset.Fields(1)
End Sub

Private Sub VSFlexGrid2_DblClick()
If Adodc8.Recordset.EOF Then Exit Sub
Adodc8.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc8.Recordset.Move rs - 1
Text1.Text = Adodc8.Recordset.Fields(1)
SSTab1.Tab = 0
End Sub

Private Sub vSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid2.RowSel
R1 = VSFlexGrid2.ColSel
End Sub

Private Sub vSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid2.RowSel
L = 0
For i = S1 To S2
L = L + Val(VSFlexGrid2.TextMatrix(i, R1))
Next
End Sub

Private Sub Command1_Click()
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "len(isnull(配料,0))<9 and 染色='N' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "len(配料)> 9 and len(染色)<9 and 染色<>'Y' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "颜色 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(Trim(DTPicker1.value) + Space(2) + Text17(0) + ":" + Text17(1) + ":" + Text17(2), "yyyy-MM-dd hh:mm:ss")
t2 = Format(Trim(DTPicker2.value) + Space(2) + Text18(0) + ":" + Text18(1) + ":" + Text18(2), "yyyy-MM-dd hh:mm:ss")
sql1 = sql1 + "cast(排产时间 as varchar(19)) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "车台='" & DataCombo6.Text & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "操作 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If Check2(12).value = 1 Then
sql1 = sql1 + "班次 like '%'+'" & Combo1.Text & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT * FROM v_kpdb where  (" + sql1 + ") ORDER BY  车台,排产编号,排产时间"
Adodc1.Refresh
Adodc6.RecordSource = "SELECT 车台,count(distinct 锅号) as 缸数,round(sum(重量),2) as 合计重量 FROM v_kpdb where  (" + sql1 + ") group by 车台 ORDER BY 车台"
Adodc6.Refresh
Adodc9.RecordSource = "SELECT count(distinct 锅号) as 缸数,round(sum(重量),2) as 合计重量 FROM v_kpdb where  (" + sql1 + ")"
Adodc9.Refresh

VSFlexGrid4.ColWidth(0) = 200
Call gssx
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
'Call pcjh(Adodc10, Adodc11, sql1)
Call MXOutadodcToExcel(VSFlexGrid4, "染缸计划" + "日期：" + Trim(DTPicker1.value) + "--" + Trim(DTPicker2.value))
End Sub

Private Sub DataCombo7_Click(Area As Integer)
DataCombo6.Text = DataCombo7.Text

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 编号 from v_ctpc_bh"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Label8.Caption = Format(Date, "yymmdd") + "-010"
Call Label4_Click
Else
uu = Val(Adodc4.Recordset.Fields(0)) + 2
Label8.Caption = Format(Date, "yymmdd") + "-" + Left("000", 3 - Len(Trim(Str(uu)))) + Trim(Str(uu))
Call Label4_Click
End If

End Sub

Private Sub Form_Load()
  Me.Move 0, 0, 8000, 6000
DataCombo1.Text = ""
DataCombo2.Text = ""
Combo1.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo6.Text = ""
DTPicker1.value = Date - 1
DTPicker2.value = Date
DTPicker3.value = Date - 1
DTPicker5.value = Date
DTPicker4.value = Now
Check2(9).value = 1
Check2(4).value = 1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text12.Text = ""
Text13.Text = ""
Text7.Text = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text14 = ""

Text15(0) = "00"
Text15(1) = "00"
Text15(2) = "01"

Text16(0) = "23"
Text16(1) = "59"
Text16(2) = "59"

Text17(0) = "00"
Text17(1) = "00"
Text17(2) = "01"

Text18(0) = "23"
Text18(1) = "59"
Text18(2) = "59"
cdbhf = cdbh
Option1.value = True
Option4.value = True
Label8.Caption = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
DataCombo9(0).Text = ""
Combo2.Text = "待染"
DataCombo10.Text = ""
DataCombo11.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset



t1 = Format(DTPicker1.value, "yyyy-mm-dd 00:00:01")
t2 = Format(DTPicker2.value, "yyyy-mm-dd 23:59:59")

jd = 1
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * FROM v_kpdb where cast(排产时间 as varchar(19)) between '" & t1 & "' and '" & t2 & "' ORDER BY 车台,排产编号,排产时间"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from khZL  group by 简称"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 车台编号 from ct  order by ip"
Adodc3.Refresh
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 编号 from v_ctpc_bh"
Adodc4.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
If Adodc4.Recordset.EOF Then
Text2.Text = Format(Date, "yymmdd") + "-011"
Else
uu = Val(Adodc4.Recordset.Fields(0)) + 2
Text2.Text = Format(Date, "yymmdd") + "-" + Left("000", 3 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select distinct 名称,序号  from cjrsfs order by 序号"
Adodc10.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct 姓名,编号  from works1 order by 编号"
Adodc11.Refresh
End Sub

Private Sub Label10_Click()
If DataCombo8.Text = "" Then
FormJ8.DataCombo4 = Text1.Text
FormJ8.Show
Else
FormJ8.DataCombo6 = DataCombo8
FormJ8.Check2(7) = 1
FormJ8.Check2(9) = 1
FormJ8.Show
End If
End Sub

Private Sub Label13_Click()
If DataCombo8.Text = "" Then
MsgBox ("请选择车台！")
Exit Sub
End If

If Text1.Text = "" Then
MsgBox ("请选择缸号！")
Exit Sub
End If

If Len(Text4.Text) < 9 Then
MsgBox ("请输入转缸编号！")
Exit Sub
End If

If Combo1.Text = "" Then
MsgBox ("请输入班次！")
Exit Sub
End If

If MsgBox("确定转缸吗？", vbYesNo) = vbNo Then Exit Sub
t1 = Format(DTPicker4.value, "yyyy-mm-dd hh:mm:ss")
sql1 = "UPDATE KPD SET ye=convert(nvarchar ,'" & t1 & "',120),rs='',车台='" & DataCombo8 & "',bh='" & Text4 & "',bz='" & Text3.Text & "',zt='染缸安排',版号='" & Text14 & "',班次='" & Combo1 & "' WHERE 锅号='" & Text1.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

MsgBox ("转缸成功！")
Call Label5_Click
Call Command1_Click
Text4 = ""
DataCombo8 = ""
End Sub

Private Sub Label2_Click()
Formd332.Text1 = Text1.Text
Formd332.Show
End Sub

Private Sub Label3_Click()
If DataCombo7.Text = "" Then
MsgBox ("请选择车台！")
Exit Sub
End If

If Text1.Text = "" Then
MsgBox ("请选择缸号！")
Exit Sub
End If

If Len(Text2.Text) <> 10 Then
MsgBox ("请输入排产编号！")
Exit Sub
End If

'If Text14.Text = "" Then
'MsgBox ("请输入染缸操作！")
'Exit Sub
'End If

'If Combo1.Text = "" Then
'MsgBox ("请输入班次！")
'Exit Sub
'End If

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "SELECT * FROM v_kpdb  WHERE 排产编号='" & Text2.Text & "' and 车台='" & DataCombo7.Text & "'"
Adodc7.Refresh
If Not Adodc7.Recordset.EOF Then
If MsgBox("已有此排产编号，是否并缸？", vbYesNo) = vbNo Then Exit Sub
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "SELECT * FROM v_kpdb  WHERE 锅号='" & Text1.Text & "' and len(排产时间)>10"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
t1 = Format(DTPicker4.value, "yyyy-mm-dd hh:mm:ss")
If Option4.value = True Then
sql1 = "UPDATE KPD SET ye=convert(nvarchar ,'" & t1 & "',120),rs='',车台='" & DataCombo7.Text & "',bh='" & Text2.Text & "',bz='" & Text3.Text & "',zt='染缸安排',版号='" & Text14 & "',cky='完',班次='" & Combo1 & "' WHERE 锅号='" & Text1.Text & "'"
Else
sql1 = "UPDATE KPD SET ye=convert(nvarchar ,'" & t1 & "',120),rs='',车台='" & DataCombo7.Text & "',bh='" & Text2.Text & "',bz='" & Text3.Text & "',zt='染缸安排',版号='" & Text14 & "',cky='续',班次='" & Combo1 & "' WHERE 锅号='" & Text1.Text & "'"
End If
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("排缸成功！")
Call Label5_Click
Else
MsgBox ("已排产")
End If
Call Command5_Click
Call Command1_Click
End Sub

Private Sub Label4_Click()
If DataCombo7.Text = "" Then
'MsgBox ("请选择车台！")
Exit Sub
End If
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 编号 from v_ctpc_bh"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Text2.Text = Format(Date, "yymmdd") + "-001"
Else
uu = Val(Adodc4.Recordset.Fields(0)) + 2
Text2.Text = Format(Date, "yymmdd") + "-" + Left("000", 3 - Len(Trim(Str(uu)))) + Trim(Str(uu))
End If
End Sub

Private Sub Label5_Click()
On Error Resume Next
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "len(isnull(配料,0))<9 and 染色='N' and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo1.Text & "'+'%' and "
End If

If Check2(2).value = 1 Then
sql1 = sql1 + "len(配料)> 9 and len(染色)<9 and 染色<>'Y' and "
End If

If Check2(3).value = 1 Then
sql1 = sql1 + "标签 like '%'+'" & DataCombo2.Text & "'+'%' and "
End If

If Check2(4).value = 1 Then
t1 = Format(DTPicker1.value, "yyyy-mm-dd 12:00:01")
t2 = Format(DTPicker2.value, "yyyy-mm-dd 11:59:59")
sql1 = sql1 + "convert(varchar(100),排产时间,120) between '" & t1 & "' and '" & t2 & "' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & DataCombo4.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "色别 like '%'+'" & DataCombo5.Text & "'+'%' and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT * FROM v_kpdb where (" + sql1 + ") and 车台='" & DataCombo7.Text & "' ORDER BY 车台,排产编号,排产时间"
Adodc1.Refresh
Adodc6.RecordSource = "SELECT 车台,count(distinct 锅号) as 缸数,sum(重量) as 合计重量 FROM v_kpdb where (" + sql1 + ") and 车台='" & DataCombo7.Text & "' group by 车台 ORDER BY 车台"
Adodc6.Refresh
Call gssx
End Sub

Private Sub Text1_Change()
If InStr(Text1.Text, "J") Then
Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
End If
End Sub

Private Sub Timer1_Timer()
If Option1.value = True Then
If jd = 2 Then
DTPicker4.value = Now
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 编号 from v_ctpc_bh"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Label8.Caption = Format(Date, "yymmdd") + "-011"
Call Label4_Click
Else
uu = Val(Adodc4.Recordset.Fields(0)) + 2
Label8.Caption = Format(Date, "yymmdd") + "-" + Left("000", 3 - Len(Trim(Str(uu)))) + Trim(Str(uu))
Call Label4_Click
End If
If DataCombo7.Text = "" Then Label8.Caption = ""
jd = 1
Else
jd = jd + 1
End If
End If
End Sub

Private Sub VSFlexGrid4_Click()
On Error Resume Next
If Option2.value = True Then
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid4.Row
Adodc1.Recordset.Move rs - 1
Text1 = Adodc1.Recordset.Fields(8)
Text4 = Adodc1.Recordset.Fields(2)
Text3 = Adodc1.Recordset.Fields(9)
Text14 = Adodc1.Recordset.Fields(10)
End If
End Sub

Private Sub VSFlexGrid4_dblClick()
'If Adodc1.Recordset.EOF Then Exit Sub
'Adodc1.Recordset.MoveFirst
'rs = VSFlexGrid4.Row
'Adodc1.Recordset.Move rs - 1
'Formd331.Text5 = Adodc1.Recordset.Fields(5)
'Formd331.Show
End Sub

Private Sub gssx()
If VSFlexGrid4.Rows > 1 Then
For i = 1 To VSFlexGrid4.Rows - 1
VSFlexGrid4.RowHeight(i) = 400
Next
End If
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If
End Sub

