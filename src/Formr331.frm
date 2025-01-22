VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Formr331 
   BackColor       =   &H00C0E0FF&
   Caption         =   "粉体半自动称量系统"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   DrawWidth       =   4684
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   19500
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "称量信息"
      TabPicture(0)   =   "Formr331.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "配料信息"
      TabPicture(1)   =   "Formr331.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Index           =   1
         Left            =   0
         ScaleHeight     =   10275
         ScaleWidth      =   18315
         TabIndex        =   99
         Top             =   600
         Width           =   18375
         Begin VB.TextBox Text14 
            Height          =   2535
            Left            =   15720
            TabIndex        =   176
            Text            =   "Text14"
            Top             =   5520
            Width           =   2415
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   24
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   15960
            TabIndex        =   174
            Text            =   "Text13"
            Top             =   3240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Timer Timer8 
            Interval        =   500
            Left            =   9360
            Top             =   240
         End
         Begin VB.Timer Timer7 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   8880
            Top             =   240
         End
         Begin VB.Timer Timer6 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   8400
            Top             =   240
         End
         Begin VB.TextBox Text11 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   24
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Index           =   4
            Left            =   13560
            TabIndex        =   108
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "关闭串口"
            Height          =   495
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   7920
            Top             =   240
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   6960
            Top             =   240
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   72
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1455
            Index           =   3
            Left            =   9840
            TabIndex        =   106
            Top             =   6600
            Width           =   5415
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   72
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1335
            Index           =   2
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   105
            Top             =   4440
            Width           =   5655
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   24
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   735
            Index           =   1
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   855
            Index           =   0
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   103
            Top             =   2640
            Width           =   5655
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "退出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   2040
            TabIndex        =   101
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   2040
            TabIndex        =   100
            Top             =   1560
            Width           =   1935
         End
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   330
            Left            =   5520
            Top             =   9840
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
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
            Left            =   6240
            Top             =   9720
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
            Left            =   5880
            Top             =   9720
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
            Left            =   5040
            Top             =   9600
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
            Height          =   375
            Left            =   5640
            Top             =   9840
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
            Height          =   375
            Left            =   6000
            Top             =   9480
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
            Top             =   9840
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
            Bindings        =   "Formr331.frx":0038
            Height          =   5535
            Left            =   240
            TabIndex        =   109
            Top             =   2640
            Width           =   8175
            _cx             =   14420
            _cy             =   9763
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
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
         Begin MSCommLib.MSComm MSComm1 
            Left            =   7080
            Top             =   720
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            CommPort        =   2
            DTREnable       =   -1  'True
            BaudRate        =   600
         End
         Begin MSCommLib.MSComm MSComm2 
            Left            =   7080
            Top             =   1320
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            CommPort        =   2
            DTREnable       =   -1  'True
            BaudRate        =   600
         End
         Begin MSCommLib.MSComm MSComm3 
            Left            =   7080
            Top             =   1920
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            CommPort        =   2
            DTREnable       =   -1  'True
            BaudRate        =   600
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFF80&
            Caption         =   "手工补料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   175
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            Caption         =   "包装取消"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   9240
            TabIndex        =   127
            Top             =   6600
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            Caption         =   "包装称重"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   8520
            TabIndex        =   126
            Top             =   6600
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            Caption         =   "继续称重"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   7
            Left            =   9240
            TabIndex        =   125
            Top             =   7440
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            Caption         =   "换筒称重"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   8
            Left            =   8520
            TabIndex        =   124
            Top             =   7440
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "称重去皮"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9720
            TabIndex        =   122
            Top             =   5880
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   9360
            TabIndex        =   120
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "提示信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   8160
            TabIndex        =   119
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "实际称重"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   8520
            TabIndex        =   118
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "需要称重"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   2
            Left            =   8520
            TabIndex        =   117
            Top             =   4440
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "染料序号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   8520
            TabIndex        =   116
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "称量染料名称"
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
            Index           =   0
            Left            =   8520
            TabIndex        =   115
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "称量信息"
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
            Index           =   1
            Left            =   240
            TabIndex        =   114
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "条码或卡号扫描"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   113
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "料单编号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   112
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "分析天平称量完成"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   11280
            TabIndex        =   111
            Top             =   3600
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFF00&
            Caption         =   "重新扫描"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   110
            Top             =   600
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   10215
         Index           =   0
         Left            =   -75000
         ScaleHeight     =   10155
         ScaleWidth      =   18435
         TabIndex        =   1
         Top             =   600
         Width           =   18495
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   2
            Left            =   8400
            TabIndex        =   130
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   1
            Left            =   7560
            TabIndex        =   129
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   0
            Left            =   6720
            TabIndex        =   128
            Top             =   1680
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   123
            Top             =   2040
            Width           =   3735
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刷新"
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
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   720
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "配料信息"
            Height          =   975
            Left            =   3120
            TabIndex        =   90
            Top             =   600
            Width           =   3255
            Begin VB.OptionButton Option1 
               BackColor       =   &H000000FF&
               Caption         =   "未称量"
               Height          =   495
               Left            =   240
               TabIndex        =   92
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H0000FF00&
               Caption         =   "已称量"
               Height          =   495
               Left            =   1680
               TabIndex        =   91
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "状态操作"
            Height          =   975
            Left            =   14040
            TabIndex        =   76
            Top             =   3360
            Visible         =   0   'False
            Width           =   2895
            Begin VB.Frame Frame9 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   345
               TabIndex        =   81
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton Option11 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "X"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option6 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Y"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   86
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option7 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "M"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   85
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Option9 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   84
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option8 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   83
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option10 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "S"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   82
                  Top             =   240
                  Width           =   495
               End
            End
            Begin VB.TextBox Text4 
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
               Left            =   840
               TabIndex        =   80
               Top             =   960
               Width           =   1575
            End
            Begin VB.CommandButton Command1 
               Caption         =   "复位"
               Height          =   420
               Left            =   1800
               TabIndex        =   79
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton Command6 
               Caption         =   "置位"
               Height          =   420
               Left            =   360
               TabIndex        =   78
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton Command7 
               Caption         =   "查询当前状态"
               Height          =   420
               Left            =   3120
               TabIndex        =   77
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Shape Shape8 
               BackColor       =   &H00FFC0C0&
               BackStyle       =   1  'Opaque
               Height          =   300
               Left            =   3720
               Shape           =   3  'Circle
               Top             =   1080
               Width           =   300
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "状态指示灯"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   67
               Left            =   2400
               TabIndex        =   89
               Top             =   1080
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "地址："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   36
               Left            =   240
               TabIndex        =   88
               Top             =   1080
               Width           =   540
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "数值操作"
            Height          =   855
            Left            =   14040
            TabIndex        =   57
            Top             =   2280
            Visible         =   0   'False
            Width           =   3255
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   960
               TabIndex        =   71
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0FFC0&
               Caption         =   "位数"
               Height          =   615
               Left            =   240
               TabIndex        =   67
               Top             =   960
               Width           =   2895
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "16位"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton Option5 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "32位"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   69
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option14 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "浮点"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   68
                  Top             =   240
                  Width           =   735
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   240
               TabIndex        =   63
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "D"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Option12 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   65
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option13 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   64
                  Top             =   240
                  Width           =   495
               End
            End
            Begin VB.TextBox Text8 
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
               Left            =   3840
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   2160
               Width           =   1575
            End
            Begin VB.CommandButton Command8 
               Caption         =   "读值"
               Height          =   420
               Left            =   3840
               TabIndex        =   61
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton Command9 
               Caption         =   "写入"
               Height          =   420
               Left            =   4440
               TabIndex        =   60
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   960
               TabIndex        =   59
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox Text7 
               Height          =   390
               Left            =   3840
               TabIndex        =   58
               Top             =   1680
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "读地址："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   51
               Left            =   360
               TabIndex        =   75
               Top             =   2160
               Width           =   720
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "输入写入数值："
               ForeColor       =   &H000040C0&
               Height          =   300
               Index           =   1
               Left            =   2280
               TabIndex        =   74
               Top             =   1800
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "写地址："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   65
               Left            =   360
               TabIndex        =   73
               Top             =   1800
               Width           =   720
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "显示读出数值："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   0
               Left            =   2280
               TabIndex        =   72
               Top             =   2160
               Width           =   1260
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "实时读Y0--Y7"
            Height          =   1815
            Index           =   0
            Left            =   11400
            TabIndex        =   10
            Top             =   4560
            Visible         =   0   'False
            Width           =   6495
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   7
               Left            =   2640
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   34
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   7
               Left            =   5640
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   33
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   6
               Left            =   2280
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   32
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   5
               Left            =   1920
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   31
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   4
               Left            =   1560
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   30
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   3
               Left            =   1200
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   29
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   2
               Left            =   840
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   28
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   1
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   27
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   26
               Top             =   1320
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   6
               Left            =   5280
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   25
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   5
               Left            =   4920
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   24
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   4
               Left            =   4560
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   23
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   3
               Left            =   4200
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   22
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   2
               Left            =   3840
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   21
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   1
               Left            =   3480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   20
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   3120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   19
               Top             =   480
               Width           =   255
            End
            Begin VB.Timer Timer3 
               Enabled         =   0   'False
               Interval        =   10
               Left            =   6000
               Top             =   840
            End
            Begin VB.Timer Timer4 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   6000
               Top             =   360
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   18
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   2
               Left            =   840
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   17
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   3
               Left            =   1200
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   16
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   4
               Left            =   1560
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   15
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   5
               Left            =   1920
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   14
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   6
               Left            =   2280
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   13
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   7
               Left            =   2640
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   12
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   8
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   11
               Top             =   480
               Width           =   255
            End
            Begin MSCommLib.MSComm MSComm4 
               Left            =   5280
               Top             =   960
               _ExtentX        =   1005
               _ExtentY        =   1005
               _Version        =   393216
               DTREnable       =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y21"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   66
               Left            =   2280
               TabIndex        =   56
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y20"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   64
               Left            =   1920
               TabIndex        =   55
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y19"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   63
               Left            =   1560
               TabIndex        =   54
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y18"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   62
               Left            =   1200
               TabIndex        =   53
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y17"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   61
               Left            =   840
               TabIndex        =   52
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y16"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   60
               Left            =   480
               TabIndex        =   51
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y15"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   59
               Left            =   120
               TabIndex        =   50
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y14"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   58
               Left            =   5280
               TabIndex        =   49
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y13"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   57
               Left            =   4920
               TabIndex        =   48
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y12"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   56
               Left            =   4560
               TabIndex        =   47
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y11"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   55
               Left            =   4200
               TabIndex        =   46
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y10"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   54
               Left            =   3840
               TabIndex        =   45
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y9"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   53
               Left            =   3480
               TabIndex        =   44
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y8"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   52
               Left            =   3120
               TabIndex        =   43
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y0"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   20
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y1"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   21
               Left            =   480
               TabIndex        =   41
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y2"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   22
               Left            =   840
               TabIndex        =   40
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y3"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   23
               Left            =   1200
               TabIndex        =   39
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y4"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   24
               Left            =   1560
               TabIndex        =   38
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y5"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   25
               Left            =   1920
               TabIndex        =   37
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y6"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   26
               Left            =   2280
               TabIndex        =   36
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y7"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   27
               Left            =   2640
               TabIndex        =   35
               Top             =   240
               Width           =   210
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "通讯口操作："
            Height          =   1335
            Index           =   0
            Left            =   8160
            TabIndex        =   2
            Top             =   240
            Width           =   6135
            Begin VB.ComboBox Combo1 
               Height          =   300
               ItemData        =   "Formr331.frx":004D
               Left            =   240
               List            =   "Formr331.frx":004F
               TabIndex        =   6
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H00C0C0FF&
               Caption         =   "打开通讯"
               Height          =   375
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command11 
               BackColor       =   &H00C0C0FF&
               Caption         =   "关闭通讯"
               Height          =   375
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   2880
               TabIndex        =   3
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "端口号："
               Height          =   180
               Left            =   240
               TabIndex        =   9
               Top             =   300
               Width           =   720
            End
            Begin VB.Label Label244 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "通讯状态："
               ForeColor       =   &H00000040&
               Height          =   300
               Index           =   1
               Left            =   1920
               TabIndex        =   8
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "请打开串口"
               ForeColor       =   &H00000040&
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   7
               Top             =   945
               Width           =   900
            End
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1680
            TabIndex        =   94
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   255
            Format          =   309198849
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   95
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   309198849
            CurrentDate     =   36892
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFFC0&
            Caption         =   "手工确认"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   11880
            TabIndex        =   173
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12312312312312323"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   885
            Index           =   0
            Left            =   9600
            TabIndex        =   172
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   1
            Left            =   9600
            TabIndex        =   171
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   2
            Left            =   9600
            TabIndex        =   170
            Top             =   6000
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   3
            Left            =   9600
            TabIndex        =   169
            Top             =   7320
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   4
            Left            =   9600
            TabIndex        =   168
            Top             =   8640
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   855
            Index           =   5
            Left            =   11760
            TabIndex        =   167
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   6
            Left            =   11760
            TabIndex        =   166
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   7
            Left            =   11760
            TabIndex        =   165
            Top             =   6000
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   8
            Left            =   11760
            TabIndex        =   164
            Top             =   7320
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   975
            Index           =   9
            Left            =   11760
            TabIndex        =   163
            Top             =   8640
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "      机台料单信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   9600
            TabIndex        =   162
            Top             =   2760
            Width           =   3975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "配料机台信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   480
            TabIndex        =   161
            Top             =   2760
            Width           =   8655
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   480
            TabIndex        =   160
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   480
            TabIndex        =   159
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   480
            TabIndex        =   158
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   480
            TabIndex        =   157
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   480
            TabIndex        =   156
            Top             =   7920
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   480
            TabIndex        =   155
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   6
            Left            =   2280
            TabIndex        =   154
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   7
            Left            =   2280
            TabIndex        =   153
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   2280
            TabIndex        =   152
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   9
            Left            =   2280
            TabIndex        =   151
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   10
            Left            =   2280
            TabIndex        =   150
            Top             =   7920
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   11
            Left            =   2280
            TabIndex        =   149
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   12
            Left            =   4080
            TabIndex        =   148
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   13
            Left            =   4080
            TabIndex        =   147
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   14
            Left            =   4080
            TabIndex        =   146
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   15
            Left            =   4080
            TabIndex        =   145
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   16
            Left            =   4080
            TabIndex        =   144
            Top             =   7920
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   17
            Left            =   4080
            TabIndex        =   143
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   18
            Left            =   5880
            TabIndex        =   142
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   19
            Left            =   5880
            TabIndex        =   141
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   20
            Left            =   5880
            TabIndex        =   140
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   21
            Left            =   5880
            TabIndex        =   139
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   22
            Left            =   5880
            TabIndex        =   138
            Top             =   7920
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   23
            Left            =   5880
            TabIndex        =   137
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   24
            Left            =   7680
            TabIndex        =   136
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   25
            Left            =   7680
            TabIndex        =   135
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   26
            Left            =   7680
            TabIndex        =   134
            Top             =   5760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   27
            Left            =   7680
            TabIndex        =   133
            Top             =   6840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   28
            Left            =   7680
            TabIndex        =   132
            Top             =   8040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   29
            Left            =   7680
            TabIndex        =   131
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000C0C0&
            Caption         =   "结束日期"
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
            Index           =   1
            Left            =   480
            TabIndex        =   98
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000C0C0&
            Caption         =   "起始日期"
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
            Index           =   1
            Left            =   480
            TabIndex        =   97
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "配单信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   96
            Top             =   2040
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "Formr331"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim a As String
Dim flag1 As Integer
Dim flag2 As Boolean
Dim flag3 As Boolean     ''''''''染料判断变量
Dim i
Dim ksjs As Integer      '''''称重稳定计数
Dim qpys  As Integer    '''''去皮延时
'''''''''''''''''             PLC 变量
Dim YMSCT As String '位元件操作选择标志
Dim Adree As String ' 元件地址
Dim Order As Integer '通讯顺序
Dim RWorder As Integer ' 读写通讯顺序
Dim RWcomm As Boolean '读取操作
Dim ysbc As Integer '''''''寄存器延时保持
Dim SJPD As Integer
Dim dqdz As Integer ''''''''判断是否数据
Dim dczw1, dczw2, dczw3, dczw4, dczw5, dczw6 As Integer ''''''''判断是否有称量数据
Dim bcbl1, bcbl2, bcbl3 As Integer ''''''''数据保存
Dim xrld, xrld1, xrld2, xrld3 As Integer ''''''''写入料单信息
Dim SBBH As Integer    '''''设备编号
Dim d1 As Integer  ''''''d1的数值
Dim dzdq(3) As String  ''''电子称变量和判断染料编号
Dim dzbl(4) As String  ''''给电子称传输变量
Dim dzdqpd As Integer  ''''那个电子称
Dim bzzl As Integer    '''包装重量
Dim htzl As Integer  ''''换筒变量
Dim sbqh As String ''''设备区号
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   '浮点数处理
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim MXH  As Integer    '''''''''循环读M
Dim gdbl As Integer ''''''关闭端口 打开端口变量
Private Sub Command1_Click()    '''元件复位
  Adree = YMSCT & Text4.Text
  a = gk528SetDevice(Adree, 0)  '地址  置位为1 复位为0
  RWorder = 8
  RWcomm = True
End Sub

Private Sub Command10_Click()
  Dim b As String
  Dim COM1 As Integer
  
  COM1 = Combo1.ListIndex + 1
  b = OpenComm(MSComm4, COM1, "9600,e,7,1")
  If b = 0 Then
     Label2(4).Caption = "串口已打开"
     Order = 0
     Timer3.Enabled = True
     Timer4.Enabled = True
     RWcomm = False
 Else
     Label2(4).Caption = "串口关闭"
     Timer4.Enabled = False
     Timer3.Enabled = False
End If

End Sub

Private Sub Command11_Click()
 Dim b As String
 b = CloseComm(MSComm4)
 Timer3.Enabled = False
 Timer4.Enabled = False
 Label2(4).Caption = "串口关闭"
End Sub

Private Sub Command2_Click()
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
        End If
If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
        End If
Timer1.Enabled = True
flag2 = False
Unload Me
End Sub

Private Sub Command3_Click()
Text9 = ""
Text3 = ""
Text9.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct isnull(机台,'') as 机台 FROM v_pldr_ft WHERE 染化助库 not like '%助剂%' and (称量标记='N' or 称量标记 is null) AND cast(CONVERT(varchar(120),配料日期,23) as datetime)  between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 染化助名称 in(select 粉体名称 from ftsb) ORDER BY 机台"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct isnull(机台,'') as 机台 FROM v_pldr_ft WHERE 染化助库 not like '%助剂%' and 称量标记='Y' AND cast(CONVERT(varchar(120),配料日期,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 染化助名称 in(select 粉体名称 from ftsb) ORDER BY 机台"
Adodc1.Refresh
End If

For i = 0 To 29
Label11(i).Visible = False
Next
If Not Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
L = 0
Do While Not Adodc1.Recordset.EOF
Label11(L).Caption = Adodc1.Recordset.Fields(0)
Label11(L).Visible = True
Adodc1.Recordset.MoveNext
L = L + 1
Loop
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(2) = 2500
End Sub



Private Sub Command5_Click()
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
        End If
If MSComm3.PortOpen = True Then
            MSComm3.PortOpen = False
        End If
Timer1.Enabled = False
Timer2.Enabled = False
End Sub



Private Sub Command6_Click()  ''''元件置位
  Adree = YMSCT & Text4.Text
  a = gk528SetDevice(Adree, 1)  '地址  置位为1 复位为0
  RWorder = 7
  RWcomm = True

End Sub

Private Sub Command7_Click()    '''查询元件状态
 Adree = YMSCT & Text4.Text
 a = gk528ReadDevice(Adree, 1)  '地址  个数
 RWorder = 9
 RWcomm = True
End Sub

Private Sub Command8_Click()       ''''''读元件
 If Option3.value = True Then 'D
    Adree = "D" & Text10.Text
 Else
    If Option12.value = True Then 'C
       Adree = "CN" & Text10.Text
    Else
       Adree = "TN" & Text10.Text
    End If
 End If
 If Option4.value = True Then
    a = gk528ReadDevice(Adree, 1)  '地址  个数
 Else
    a = gk528ReadDevice(Adree, 2)
 End If
 RWorder = 5
 RWcomm = True
End Sub

Private Sub Command9_Click()   '''''' 写元件
 Dim Number As String
    '写入数值
 Dim WriteData() As String
 
 If Option4.value = True Then 'D
    Adree = "D" & Text5.Text
 Else
    If Option12.value = True Then 'C
       Adree = "CN" & Text5.Text
    Else
       Adree = "TN" & Text5.Text
    End If
 End If
 
 If Option4.value = True Then '16位
    ReDim WriteData(0) As String
    WriteData(0) = Val(Text7.Text)
    a = gk528WriteDevice(Adree, 1, WriteData)   '地址  个数  数值组
 End If
 RWorder = 6
 RWcomm = True
End Sub


Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date
DTPicker2.value = Date
dzdqpd = 0   '''''初始称料状态读取
Label4.Caption = ""
gdbl = 0
Text3 = ""
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Text13 = 0
MSComm2.CommPort = 2
MSComm2.Settings = "600,n,8,1"

MSComm3.CommPort = 3
MSComm3.Settings = "600,n,8,1"

flag1 = 4 ''''''''不显示称重量

flag2 = True
flag3 = False
For m = 0 To 4
Text1(m) = ""
Text12(m) = ""
Next

Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Text14 = ""
bzzl = 0

For i = 0 To 9
Label12(i).Visible = False
Next
For i = 0 To 29
Label11(i).Visible = False
Next


  Dim b As String
  
  b = OpenComm(MSComm4, 1, "9600,e,7,1")
  
  If b = 0 Then
     Label2(4).Caption = "串口已打开"
     Order = 0
     Timer3.Enabled = True
     Timer4.Enabled = True
     RWcomm = False
 Else
     Label2(4).Caption = "串口关闭"
     Timer4.Enabled = False
     Timer3.Enabled = False
 End If


    Dim g As Integer
      '*添加通讯口选择变量
      
    For g = 1 To 10                             '*添加通讯口选择
        Combo1.AddItem "Com" & CStr(g)
    Next g
    Combo1.ListIndex = 0  '显示第一项
    Option7.value = True
    YMSCT = "M"
    DCT = "D"



Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库 not like '%助剂%' and (称量标记='N' or 称量标记 is null) AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库 not like '%助剂%' and 称量标记='Y' AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Text2.TabIndex = 0
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(3) = 2300


End Sub



Private Sub Label10_Click()
Text3 = ""
Text2 = ""
Text2.SetFocus
End Sub

Private Sub Label11_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct 料单编号,锅号 FROM v_pldr_ft WHERE 染化助库 not like '%助剂%' and (称量标记='N' or 称量标记 is null) AND cast(CONVERT(varchar(120),配料日期,23) as datetime)  between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 染化助名称 in(select 粉体名称 from ftsb) and isnull(机台,'')='" & Label11(Index).Caption & "'  ORDER BY 料单编号 desc"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 料单编号,锅号 FROM v_pldr_ft WHERE 染化助库 not like '%助剂%' and 称量标记='Y' AND cast(CONVERT(varchar(120),配料日期,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and 染化助名称 in(select 粉体名称 from ftsb) and isnull(机台,'')='" & Label11(Index).Caption & "'  ORDER BY 料单编号 desc"
Adodc1.Refresh
End If

For i = 0 To 9
Label12(i).Visible = False
Next
If Not Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
L = 0
Do While Not Adodc1.Recordset.EOF
Label12(L).Caption = Adodc1.Recordset.Fields(0)
Label12(L).Visible = True
Adodc1.Recordset.MoveNext
L = L + 1
Loop
End If
      End Select
End Sub

Private Sub Label12_Click(Index As Integer)
Select Case Index
       Case Index
       Text3 = Label12(Index).Caption
       SSTab1.Tab = 0
End Select
End Sub

Private Sub Label13_Click()
Formr452.Show
End Sub

Private Sub Label14_Click()
On Error Resume Next
If MsgBox("确定手工确认吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "UPDATE pldr SET 实际称量=配料用量,称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
lnl = Text3
Text3 = ""
Text3 = lnl
End Sub

Private Sub Label15_Click()
On Error Resume Next
If dzdqpd = 2 Then
 If MSComm3.PortOpen = False Then
            MSComm3.PortOpen = True
      If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''经典没有端口就退出
 End If
   MSComm3.Output = Chr$(27) + "t"
End If
 
If dzdqpd = 1 Then
 If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
            If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''经典没有端口就退出
 End If
   MSComm2.Output = Chr$(27) + "t"
End If

If dzdqpd = 4 Then
 If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
            If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''经典没有端口就退出
 End If
   MSComm2.Output = Chr$(27) + "t"
End If

If dzdqpd = 3 Then
 If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
            If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''经典没有端口就退出
 End If
   MSComm1.Output = Chr$(27) + "t"
End If
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 3
pmbl = 1
Formr440.Text1 = Text3
Formr440.Show
End Select
End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case 5
bzzl = 0
Text1(2) = Val(Text13)
Text13 = 0
       Case 6
Adodc5.RecordSource = "select 包装数量 from ftsb where 粉体名称='" & Text1(0) & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
'Call Label15_Click
bzzl = Val(Adodc5.Recordset.Fields(0)) * 1000
Text13 = Text1(2)
Text1(2) = bzzl - Val(Text13)
Else
bzzl = 0
Text13 = 0
End If

     Case 8

cll = Format(Val(Text1(3)) / 1000, "#0.00000")   ''''''''''称量单位g转换成kg
sql1 = "UPDATE pldr SET 实际称量=(isnull(实际称量,0)+'" & cll & "') WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc2.Refresh


     Case 7

Call VQJC
End Select

End Sub

Private Sub Label8_Click()
cll = Text1(2)
sql1 = "UPDATE pldr SET 实际称量='" & cll & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc2.Refresh
Label8.Visible = False                ''''''''''关闭分析称量
Call VQJC
Call Command4_Click
Text1(0).ForeColor = &HFF&

End Sub


Private Sub MSComm4_OnComm()
On Error Resume Next
 Dim b As String
 Dim i As Integer
 Dim Tdata1 As String, Tdata2 As String, Tdata3 As String, Tdata4 As String '*临时变量
 Dim Ddata(6) As Long '*中间变量
 Dim Mdata(1) As Integer '*中间变量
                      Dim Data10 As Long    '*浮点数中间处理变量；
                      Dim Data As Single    '*浮点数中间处理变量；
                      Dim dataCl As String  '*浮点数中间处理变量；
    
   b = MSCONComm(MSComm4)
   Text6.Text = b
   If b <> "0" Then Exit Sub
   Timer4.Enabled = False
   Select Case Order
          Case 0   'read d102-105
                         
                        Ddata(0) = "&H" + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2)  '*PLC返回的寄存器数值是从低字节到高字节排列，所以我们要重新排列一下！
                        dzdq(3) = CStr(Val(Ddata(0)))
                        Text12(2) = Val(dzdq(3))
               
          Case 6, 7, 8  '写 置，复位
               Order = 0
   End Select

   Timer3.Enabled = True

End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 3
'If Val(Text1(3)) > 0 And Val(Text1(2)) <= (Val(Text1(3)) + 0.02) And Val(Text1(2)) >= (Val(Text1(3)) - 0.02) And Val(Text1(2)) > 0 Then     '''''判断是否保存
'Timer2.Enabled = True
'ksjs = 0
'End If
       Case 4
If Text1(4) = "0" Then
Timer1.Enabled = False
End If

If Text1(4) = "1" Then
Label4.Caption = "请注意是否去皮！！"
Beep 2000, 50
qpys = 3                                ''''''''延时准备变量为20秒
Timer1.Enabled = True
Text1(4) = ""
End If
End Select
End Sub

Private Sub Text10_Change()
 If Option3.value = True Then 'D
    Adree = "D" & Text10.Text
 Else
    If Option12.value = True Then 'C
       Adree = "CN" & Text10.Text
    Else
       Adree = "TN" & Text10.Text
    End If
 End If
 If Option4.value = True Then
    a = gk528ReadDevice(Adree, 1)  '地址  个数
 Else
    a = gk528ReadDevice(Adree, 2)
 End If
 RWorder = 5
 RWcomm = True
End Sub

Private Sub Text2_Change()
If InStr(Text2, "J") > 0 Or InStr(Text2, "j") > 0 Then
Text3 = Mid(Text2, 1, Len(Text2) - 1)
Text2 = ""
Text2.SetFocus
End If
End Sub

Private Sub Text3_Change()
'On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call VQJC

'Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select * from FTBZXS where 料单编号='" & Text3 & "'"
'Adodc4.Refresh
'If Not Adodc4.Recordset.EOF Then
'Formr333.Text3 = Text3
'Formr333.Show
'End If
End Sub


Private Sub Text4_Change()
If Val(Text4) = 1 Then
  Adree = "M66"
  a = gk528SetDevice(Adree, 1)  '地址  置位为1 复位为0
  RWorder = 7
  RWcomm = True
End If
If Val(Text4) = 0 Then
  Adree = "M66"
  a = gk528SetDevice(Adree, 0)  '地址  置位为1 复位为0
  RWorder = 7
  RWcomm = True
End If
Text4 = ""
End Sub

Private Sub Text8_Change()
Text1(4) = Text8
End Sub



Private Sub Text9_Change()
If InStr(Text9, "J") > 0 Then
gh = Mid(Text9, 1, Len(Text9) - 1)
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库 not like '%助剂%' and (称量标记='N' or 称量标记 is null) AND 锅号='" & gh & "' ORDER BY 料单编号"
Adodc1.Refresh
Text9 = ""
End If
End Sub

Private Sub Timer1_Timer()
If qpys = 1 Then    ''''''去皮延时
Timer1.Enabled = False
End If

qpys = qpys - 1
Label4.Caption = "请注意是否去皮！！" + Trim(qpys)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Timer2_Timer()
    On Error GoTo errorhandler

    ' 校验并处理不同状态
    Select Case dzdqpd
        Case 1
            Call HandleState(0.1, 0) ' 允许误差 ±0.1g
            
        Case 4
            Call HandleState(1, 30) ' 允许误差 ±1g
            
        Case 2
            Call HandleState(1, 100) ' 允许误差 ±1g
    End Select

    ' 检查 ksjs 是否达到 2，完成称量
    If ksjs = 2 Then
        Timer2.Enabled = False
        flag1 = 4

        Dim cll As String
        If Val(Text13) = 0 Then
            cll = Format(Val(Text1(3)) / 1000, "#0.00000") ' 转换成 kg
        Else
            cll = Format((bzzl - Val(Text1(3))) / 1000, "#0.00000") ' 转换成 kg
        End If

        ' 更新数据库
        sql1 = "UPDATE pldr SET 实际称量='" & cll & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
        RD.Open sql1, conn, adOpenStatic, adLockOptimistic

        ' 刷新显示
        Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
        Adodc2.Refresh

        ' 清理状态
        Call ResetState

        ' 准备下一步
        qpys = 6 ' 延时准备放料筒盖子
        Timer6.Enabled = True
        Call Command4_Click
        Text1(0).ForeColor = &HFF&
    End If

    Exit Sub

errorhandler:
    MsgBox "发生错误：" & Err.Description, vbCritical
    Resume Next
End Sub

' 处理称量逻辑的通用子过程
Private Sub HandleState(allowedError As Double, minValue As Double)
    If Val(Text1(3)) > minValue And _
       (Val(Text1(2)) - Val(Text1(3))) <= allowedError And _
       (Val(Text1(2)) - Val(Text1(3))) >= -allowedError And _
       Val(Text1(2)) > minValue Then

        ksjs = ksjs + 1
        Beep 1000, 50

        ' 切换颜色
        If ksjs Mod 2 = 0 Then
            Text1(0).ForeColor = &HFF&
        Else
            Text1(0).ForeColor = &HFF00&
        End If
    Else
        ksjs = 0
        Text1(0).ForeColor = &HFF&
    End If
End Sub

' 清理状态的子过程
Private Sub ResetState()
    Text13 = 0
    bzzl = 0
    Text4 = 0
    Text11 = ""

    dzbl(1) = 0
    dzbl(2) = 0
    dzbl(3) = 0
    dzbl(4) = 0
    dzdqpd = 0
    bzzl = 0
    ksjs = 0

    ' 清空写入数据
    ReDim WriteData(0 To 14) As String

    ' 写入设备
    Dim a As Integer
    a = gk528WriteDevice("D100", 2, WriteData())
    RWorder = 6
    RWcomm = True
End Sub


Private Sub VQJC()
On Error Resume Next
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT ISNULL(称量标记,'N'),工序名称,染化助库,染化助名称,配料单位,round(配料用量,6),isnull(实际称量,0),次序号,包装数量,设备编号,isnull(设备区位,'0') FROM v_pldr_ft WHERE (称量标记<>'Y' OR 称量标记 IS NULL) AND 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Text11 = ""
Label4.Caption = "称重完成"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''称量后置位
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from ftbzxs where 料单编号='" & Text3 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Formr333.Text3 = Text3
Formr333.Show
End If
sbqh = ""
SBBH = 0   ''''''''''''''''''''        设备编号
dzdqpd = 0
dzbl(1) = 0
dzbl(2) = 0
dzbl(3) = 0
dzbl(4) = 0
Timer7.Enabled = True
Text3 = ""
Text2.SetFocus
Else
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''根据称量转换串口
If Adodc3.Recordset.Fields(0) <> "Y" Then
sbqh = Adodc3.Recordset.Fields(10)                      ''''''设备区号
Text1(0) = Adodc3.Recordset.Fields(3)
Text1(1) = Adodc3.Recordset.Fields(7)
''''''''''''''''''''''''''''''''''''''''''''''''''''''判断是否有整包装数量
If (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) >= Adodc3.Recordset.Fields(8) Then
bzsl = Int((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) / Adodc3.Recordset.Fields(8))    '''''取包装箱数
Text1(2) = (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6) - bzsl * Adodc3.Recordset.Fields(8)) * 1000  '''''转换g
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''保存包装箱数
sql1 = "delete from FTBZXS where 料单编号='" & Text3 & "' and 粉体名称='" & Text1(0) & "'"
sql2 = "insert into FTBZXS(料单编号,粉体名称,包装箱数) VALUES('" & Text3 & "','" & Text1(0) & "','" & bzsl & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Else
Text1(2) = (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 '''''转换g
End If


If Val(Text1(2)) < 0.1 And Val(Text1(2)) > 0 Then
MsgBox ("请用分析天平称量")
Label8.Visible = True
Text4 = 0
Exit Sub
End If


If Val(Text1(2)) = 0 Then
flag1 = 0
flag3 = False
Timer1.Enabled = False
Timer7.Enabled = True
Timer2.Enabled = False
sql1 = "UPDATE pldr SET 实际称量=0,称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
qpys = 10                                ''''''''延时准备放料筒盖子原为20秒
Timer6.Enabled = True
End If


VSFlexGrid2.ColFormat(5) = "#0.####"
VSFlexGrid2.ColFormat(6) = "#0.####"

If Val(Text1(2)) <= 250 And Val(Text1(2)) >= 0.1 And sbqh <> "3" Then
Text1(2) = Format(Text1(2), "#0.00")
flag1 = 0
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在1号称放入容器！"
'''''''''''''''''''''''''''''''''''''''''''''''
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 1
dzbl(1) = Adodc3.Recordset.Fields(9)
dzbl(2) = 1

Timer7.Enabled = True
Timer2.Enabled = True
''''''''''''''''''''''''''''''''''''''''''''''
Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If

'If Val(Text1(2)) <= 100 And Val(Text1(2)) > 250 And sbqh <> "3" Then
'Text1(2) = Format(Text1(2), "#0.0")
'flag1 = 0
'flag3 = False
'Timer1.Enabled = False
'Label4.Caption = "请在2号称放入容器！"
'''''''''''''''''''''''''''''''''''''''''''''''
'SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
'dzdqpd = 4
'dzbl(1) = Adodc3.Recordset.Fields(9)
'dzbl(2) = 1

'Timer7.Enabled = True
'Timer2.Enabled = True
''''''''''''''''''''''''''''''''''''''''''''''
'Text1(3) = 0
'Text1(4) = ""
'Text1(4).SetFocus
'End If


If Val(Text1(2)) > 250 And sbqh <> "3" Then
Text1(2) = Format(Text1(2), "#0")
flag1 = 1
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在2号称放入容器！"
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 2
dzbl(1) = Adodc3.Recordset.Fields(9)
dzbl(2) = 1

Timer7.Enabled = True
Timer2.Enabled = True

Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If

If Val(Text1(2)) > 250 And sbqh = "3" Then
Text1(2) = Format(Text1(2), "#0")
flag1 = 1
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在2号称放入容器！"
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 3
dzbl(1) = Adodc3.Recordset.Fields(9)
dzbl(2) = 1

Timer7.Enabled = True
Timer2.Enabled = True

Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If

Exit Sub
End If


Adodc3.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Timer3_Timer()    ''''''''''''''PLC

 If RWcomm = True Then
   Order = RWorder
   RWcomm = False
 End If
  Select Case Order
         Case 0   '读D704
              a = gk528ReadDevice("D100", 1)
 End Select
 

 MSComm4.OutBufferCount = 0 '*设置并返回发送缓冲区的字节数,设为0时清空发送缓冲区
 MSComm4.InBufferCount = 0  '*设置并返回接收缓冲区的字节数,设为0时清空接收缓冲区
 PLCText = ""
 If a = "0" Then MSComm4.Output = SenData
 Timer3.Enabled = False
 Timer4.Enabled = True

End Sub

Private Sub Timer4_Timer()              ''''plc

 If MSComm4.PortOpen = True Then
   Timer3.Enabled = True
   RWcomm = False
   Order = 0
 Else
    Timer3.Enabled = False
 End If

End Sub


Private Sub Timer5_Timer()

End Sub

Private Sub Timer6_Timer()
If qpys <= 0 Then    ''''''去皮延时
Timer6.Enabled = False
Call VQJC
End If
qpys = qpys - 1
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
If SBBH <> dzdq(3) Then
       ReDim WriteData(0 To 14) As String  ''''''写入个数
       Dim DataW As String    '*浮点数的中间处理变量；
       Dim Data10(7) As Single   '*浮点数的中间处理变量；
       Dim Buffer(3) As Byte   '*浮点数的中间处理变量；
 
       For i = 0 To 1
       WriteData(i) = Val(dzbl(i + 1))
       Next
       a = gk528WriteDevice("D100", 2, WriteData())
 RWorder = 6
 RWcomm = True
Else
Timer7.Enabled = False
End If
End Sub


Private Sub VSFlexGrid1_Click()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
Text3 = Adodc1.Recordset.Fields(2)
End Sub

Private Sub Timer8_Timer()
    On Error GoTo errorhandler ' 发生错误时跳转到errorhandler

    If dzdqpd = 0 Then ' 判断dzdqpd的值是否为0
        ClosePorts ' 关闭所有端口
        Label4.Caption = "称料完成！" ' 设置Label4的标题为“称料完成！”
    ElseIf dzdqpd = 1 Then ' 判断dzdqpd的值是否为1
        Label4.Caption = "请在1号称放入容器！" ' 设置Label4的标题为“请在1号称放入容器！”
        OpenPort MSComm2
        MSComm2.Output = Chr$(27) + "p" ' 发送命令到MSComm2端口以获取称重量
        ReadPortData MSComm2, False ' 读取MSComm2端口的输入，不需要乘以1000
    ElseIf dzdqpd = 2 Or dzdqpd = 4 Then ' 判断dzdqpd的值是否为2或4
        Label4.Caption = "请在2号称放入容器！" ' 设置Label4的标题为“请在2号称放入容器！”
        OpenPort MSComm3
        MSComm3.Output = Chr$(27) + "p" ' 发送命令到MSComm3端口以获取称重量
        ReadPortData MSComm3, True ' 读取MSComm3端口的输入，需要乘以1000
    End If

    Exit Sub ' 退出子程序

errorhandler:
    MsgBox "发生错误: " & Err.Description ' 显示错误信息
    Exit Sub
End Sub

' 子程序：关闭所有端口
Private Sub ClosePorts()
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    If MSComm2.PortOpen Then MSComm2.PortOpen = False
    If MSComm3.PortOpen Then MSComm3.PortOpen = False
End Sub

' 子程序：打开指定端口
Private Sub OpenPort(comm As MSComm)
    If Not comm.PortOpen Then comm.PortOpen = True
End Sub

' 子程序：读取端口数据
Private Sub ReadPortData(comm As MSComm, multiply As Boolean)
    On Error GoTo errorhandler
    Dim Data As String
    Dim TempData As String
    Dim CompleteData As String
    Dim retryCount As Integer
    Dim maxRetries As Integer

    ' 设置最大重试次数
    maxRetries = 5
    retryCount = 0

retry:
    ' 清空Text14内容
    Text14.Text = ""
    CompleteData = ""

    ' 等待数据到达
    Do
        DoEvents ' 处理其他事件
        If comm.InBufferCount > 0 Then
            TempData = comm.Input ' 读取当前缓冲区的数据
            CompleteData = CompleteData & TempData ' 将临时数据添加到CompleteData中
            Text14.Text = CompleteData ' 更新Text14内容
            Debug.Print "TempData: " & TempData ' 调试信息
            Debug.Print "CompleteData: " & CompleteData ' 调试信息
        End If
        
        ' 增加等待时间，确保数据到达
        Sleep 50
        
        ' 检查重试次数
        retryCount = retryCount + 1
        If retryCount > maxRetries Then Exit Do ' 超过重试次数退出循环

    Loop Until Len(CompleteData) >= 13 ' 假设数据长度为13个字符

    ' 检查数据长度是否满足要求
    'If Len(CompleteData) < 13 Then
    '    MsgBox "未能接收到完整数据"
    '    Exit Sub
   ' End If

    ' 读取最终数据
    Data = CompleteData
    Debug.Print "Final Data: " & Data ' 调试信息

    ' 根据multiply参数处理数据
    If multiply Then
        Text1(3).Text = Format(Val(Trim(Mid(Data, 1, 9))) * 1000, "#0") ' 将读取的重量值乘以1000后格式化后赋值给Text1(3)
        Debug.Print "Processed Data (multiply): " & Text1(3).Text ' 调试信息
    Else
        Text1(3).Text = Format(Val(Trim(Mid(Data, 1, 9))), "#0.00") ' 将读取的重量值格式化后赋值给Text1(3)
        Debug.Print "Processed Data: " & Text1(3).Text ' 调试信息
    End If

    Exit Sub

errorhandler:
    MsgBox "读取数据时发生错误: " & Err.Description ' 显示错误信息
    If retryCount < maxRetries Then
        retryCount = retryCount + 1
        Debug.Print "Retrying... Attempt: " & retryCount
        Resume retry
    Else
        MsgBox "多次重试后仍然无法读取数据。"
    End If
    Exit Sub
End Sub


