VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Formr442 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Һ�������Զ�����ϵͳ"
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
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   19500
      _Version        =   393216
      Tab             =   1
      TabHeight       =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "������Ϣ"
      TabPicture(0)   =   "Formr442.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "������Ϣ"
      TabPicture(1)   =   "Formr442.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�ϵ���Ϣ"
      TabPicture(2)   =   "Formr442.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2(0)"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Index           =   1
         Left            =   -75000
         ScaleHeight     =   10275
         ScaleWidth      =   16995
         TabIndex        =   148
         Top             =   720
         Width           =   17055
         Begin VB.TextBox Text12 
            Height          =   495
            Left            =   1440
            TabIndex        =   172
            Text            =   "Text2"
            Top             =   3960
            Width           =   3135
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "���ͷ�ʽ"
            Height          =   1455
            Left            =   240
            TabIndex        =   169
            Top             =   120
            Width           =   4335
            Begin VB.OptionButton Option16 
               BackColor       =   &H00C0C000&
               Caption         =   "ֻ����"
               Height          =   615
               Left            =   360
               TabIndex        =   171
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton Option17 
               BackColor       =   &H00C0C000&
               Caption         =   "����������"
               Height          =   615
               Left            =   2400
               TabIndex        =   170
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "ѡ��������"
            Height          =   3735
            Left            =   4800
            TabIndex        =   162
            Top             =   120
            Width           =   2535
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   168
               Top             =   240
               Width           =   2175
            End
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   167
               Top             =   840
               Width           =   2175
            End
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   166
               Top             =   1440
               Width           =   2175
            End
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   165
               Top             =   2040
               Width           =   2175
            End
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   164
               Top             =   2640
               Width           =   2175
            End
            Begin VB.OptionButton Option15 
               BackColor       =   &H000080FF&
               Caption         =   "Option15"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   163
               Top             =   3240
               Width           =   2175
            End
         End
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   1440
            TabIndex        =   161
            Text            =   "Text2"
            Top             =   3240
            Width           =   3135
         End
         Begin VB.Timer Timer2 
            Interval        =   500
            Left            =   8760
            Top             =   120
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   8280
            Top             =   120
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   1440
            TabIndex        =   160
            Text            =   "Text2"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   1440
            TabIndex        =   159
            Text            =   "Text2"
            Top             =   2520
            Width           =   3135
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "�رմ���"
            Height          =   495
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Text17 
            Height          =   375
            Left            =   5400
            TabIndex        =   157
            Text            =   "Text17"
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Timer Timer5 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9240
            Top             =   120
         End
         Begin VB.Timer Timer6 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9720
            Top             =   120
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ͨѶ�ڲ�����"
            Height          =   975
            Index           =   0
            Left            =   7440
            TabIndex        =   149
            Top             =   240
            Width           =   8655
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
               ItemData        =   "Formr442.frx":0054
               Left            =   1080
               List            =   "Formr442.frx":0056
               TabIndex        =   154
               Text            =   "COM1"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H00C0C0FF&
               Caption         =   "��ͨѶ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   153
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton Command11 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�ر�ͨѶ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   5160
               Style           =   1  'Graphical
               TabIndex        =   152
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text6 
               Height          =   495
               Left            =   2520
               TabIndex        =   151
               Text            =   "Text6"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton Command2 
               BackColor       =   &H00C0C0FF&
               Caption         =   "�˳�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   6600
               Style           =   1  'Graphical
               TabIndex        =   150
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�˿ںţ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   156
               Top             =   360
               Width           =   1140
            End
            Begin VB.Label Label244 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ͨѶ��"
               ForeColor       =   &H00000040&
               Height          =   300
               Index           =   1
               Left            =   2040
               TabIndex        =   155
               Top             =   360
               Width           =   540
            End
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
            Bindings        =   "Formr442.frx":0058
            Height          =   4935
            Left            =   240
            TabIndex        =   173
            Top             =   4440
            Width           =   14535
            _cx             =   25638
            _cy             =   8705
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FormatString    =   $"Formr442.frx":006D
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
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
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
            Index           =   6
            Left            =   240
            TabIndex        =   187
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFF00&
            Caption         =   "����ɨ��"
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
            Left            =   3240
            TabIndex        =   186
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
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
            Index           =   5
            Left            =   240
            TabIndex        =   185
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1�ų�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   975
            Index           =   0
            Left            =   7440
            TabIndex        =   184
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4800
            TabIndex        =   183
            Top             =   3960
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "����򿨺�ɨ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   240
            TabIndex        =   182
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF80&
            Caption         =   "�ϵ����"
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
            Index           =   3
            Left            =   240
            TabIndex        =   181
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2�ų�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   975
            Index           =   1
            Left            =   7440
            TabIndex        =   180
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "3�ų�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Index           =   2
            Left            =   11880
            TabIndex        =   179
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "4�ų�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   975
            Index           =   3
            Left            =   11880
            TabIndex        =   178
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "����׼���С�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   9000
            TabIndex        =   177
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "����׼���С�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   9000
            TabIndex        =   176
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "����׼���С�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   3
            Left            =   13440
            TabIndex        =   175
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "����׼���С�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   4
            Left            =   13440
            TabIndex        =   174
            Top             =   2880
            Width           =   2655
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   10215
         Index           =   0
         Left            =   -75000
         ScaleHeight     =   10155
         ScaleWidth      =   18435
         TabIndex        =   54
         Top             =   720
         Width           =   18495
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ʵʱ��Y0--Y7"
            Height          =   1575
            Index           =   0
            Left            =   10800
            TabIndex        =   96
            Top             =   360
            Visible         =   0   'False
            Width           =   3375
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   8
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   120
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
               TabIndex        =   119
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
               TabIndex        =   118
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
               TabIndex        =   117
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
               TabIndex        =   116
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
               TabIndex        =   115
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
               TabIndex        =   114
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   113
               Top             =   480
               Width           =   255
            End
            Begin VB.Timer Timer4 
               Enabled         =   0   'False
               Interval        =   100
               Left            =   720
               Top             =   720
            End
            Begin VB.Timer Timer3 
               Enabled         =   0   'False
               Interval        =   10
               Left            =   240
               Top             =   720
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   3120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               TabIndex        =   109
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
               TabIndex        =   108
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
               TabIndex        =   107
               Top             =   480
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
               TabIndex        =   106
               Top             =   480
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   103
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
               TabIndex        =   102
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
               TabIndex        =   101
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
               TabIndex        =   100
               Top             =   1320
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
               TabIndex        =   99
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
               TabIndex        =   98
               Top             =   480
               Width           =   255
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   7
               Left            =   2640
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   97
               Top             =   1320
               Width           =   255
            End
            Begin MSCommLib.MSComm MSComm4 
               Left            =   1320
               Top             =   600
               _ExtentX        =   1005
               _ExtentY        =   1005
               _Version        =   393216
               DTREnable       =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y7"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   142
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y6"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   141
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y5"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   140
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y4"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   139
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y3"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   138
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y2"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   137
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y1"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   136
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y0"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   135
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y8"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   134
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y9"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   133
               Top             =   240
               Width           =   210
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y10"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   132
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y11"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   131
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y12"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   130
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y13"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   129
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y14"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   128
               Top             =   240
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y15"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   127
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y16"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   126
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y17"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   125
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y18"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   124
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y19"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   123
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y20"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   122
               Top             =   1080
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y21"
               BeginProperty Font 
                  Name            =   "����"
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
               TabIndex        =   121
               Top             =   1080
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "��ֵ����"
            Height          =   1095
            Left            =   9600
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox Text7 
               Height          =   390
               Left            =   3840
               TabIndex        =   91
               Text            =   "Text7"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   960
               TabIndex        =   90
               Text            =   "Text5"
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton Command9 
               Caption         =   "д��"
               Height          =   420
               Left            =   4440
               TabIndex        =   89
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton Command8 
               Caption         =   "��ֵ"
               Height          =   420
               Left            =   3840
               TabIndex        =   88
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox Text8 
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
               Left            =   3840
               Locked          =   -1  'True
               TabIndex        =   87
               Text            =   "text8"
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ԫ��ѡ��"
               Height          =   615
               Left            =   240
               TabIndex        =   83
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton Option13 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   86
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option12 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   85
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "D"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   84
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0FFC0&
               Caption         =   "λ��"
               Height          =   615
               Left            =   240
               TabIndex        =   79
               Top             =   960
               Width           =   2895
               Begin VB.OptionButton Option14 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "����"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   82
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option5 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "32λ"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   81
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "16λ"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   80
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   960
               TabIndex        =   78
               Text            =   "Text10"
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ʾ������ֵ��"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   0
               Left            =   2280
               TabIndex        =   95
               Top             =   2160
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "д��ַ��"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   65
               Left            =   360
               TabIndex        =   94
               Top             =   1800
               Width           =   720
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����д����ֵ��"
               ForeColor       =   &H000040C0&
               Height          =   300
               Index           =   1
               Left            =   2280
               TabIndex        =   93
               Top             =   1800
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ַ��"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   51
               Left            =   360
               TabIndex        =   92
               Top             =   2160
               Width           =   720
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "״̬����"
            Height          =   1095
            Left            =   11520
            TabIndex        =   63
            Top             =   840
            Visible         =   0   'False
            Width           =   5295
            Begin VB.CommandButton Command7 
               Caption         =   "��ѯ��ǰ״̬"
               Height          =   420
               Left            =   3120
               TabIndex        =   74
               Top             =   1560
               Width           =   1335
            End
            Begin VB.CommandButton Command6 
               Caption         =   "��λ"
               Height          =   420
               Left            =   360
               TabIndex        =   73
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "��λ"
               Height          =   420
               Left            =   1800
               TabIndex        =   72
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox Text4 
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
               Left            =   840
               TabIndex        =   71
               Text            =   "text4"
               Top             =   960
               Width           =   1575
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ԫ��ѡ��"
               Height          =   615
               Left            =   345
               TabIndex        =   64
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton Option10 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "S"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   70
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option8 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   69
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option9 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   68
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option7 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "M"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   67
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Option6 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Y"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   66
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option11 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "X"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   240
                  Width           =   495
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ַ��"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   36
               Left            =   240
               TabIndex        =   76
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "״ָ̬ʾ��"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   67
               Left            =   2400
               TabIndex        =   75
               Top             =   1080
               Width           =   900
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
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "������Ϣ"
            Height          =   975
            Left            =   3120
            TabIndex        =   60
            Top             =   600
            Width           =   3255
            Begin VB.OptionButton Option2 
               BackColor       =   &H0000C0C0&
               Caption         =   "�ѳ���"
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
               Left            =   1680
               TabIndex        =   62
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H0000C0C0&
               Caption         =   "δ����"
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
               Left            =   240
               TabIndex        =   61
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "ˢ��"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   59
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   0
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "Text11"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   1
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "Text11"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   9000
            TabIndex        =   56
            Text            =   "Text14"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   2
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "Text11"
            Top             =   1200
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1680
            TabIndex        =   143
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   255
            Format          =   402718721
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   144
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   402718721
            CurrentDate     =   36892
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "      ��̨�ϵ���Ϣ"
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
            Index           =   4
            Left            =   9960
            TabIndex        =   228
            Top             =   2040
            Width           =   3975
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12120
            TabIndex        =   227
            Top             =   7920
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12120
            TabIndex        =   226
            Top             =   6600
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12120
            TabIndex        =   225
            Top             =   5280
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12120
            TabIndex        =   224
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12120
            TabIndex        =   223
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   9960
            TabIndex        =   222
            Top             =   7920
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   9960
            TabIndex        =   221
            Top             =   6600
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   9960
            TabIndex        =   220
            Top             =   5280
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   9960
            TabIndex        =   219
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H0000C0C0&
            Caption         =   "Label12312312312312323"
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   9960
            TabIndex        =   218
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   217
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   216
            Top             =   7320
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   215
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   214
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   213
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   212
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   211
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   210
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   209
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   208
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   207
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   206
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   205
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   204
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   203
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   202
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   201
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   200
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   199
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   198
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   197
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   196
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   195
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   194
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   193
            Top             =   8280
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   192
            Top             =   7200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   191
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   190
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   189
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   188
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���ϻ�̨��Ϣ"
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
            Index           =   0
            Left            =   480
            TabIndex        =   147
            Top             =   2040
            Width           =   8655
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000C0C0&
            Caption         =   "��ʼ����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   146
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   145
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Left            =   120
         ScaleHeight     =   10275
         ScaleWidth      =   18435
         TabIndex        =   1
         Top             =   720
         Width           =   18495
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   6000
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   645
            Index           =   3
            Left            =   2160
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   4080
            Width           =   5295
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   3360
            Width           =   5295
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Index           =   1
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   2760
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Index           =   5
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "����"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Index           =   0
            Left            =   2160
            TabIndex        =   23
            Text            =   "Text13"
            Top             =   6600
            Width           =   5295
         End
         Begin VB.TextBox Text13 
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
            Index           =   1
            Left            =   2160
            TabIndex        =   22
            Text            =   "Text13"
            Top             =   7560
            Width           =   1575
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   2160
            TabIndex        =   21
            Text            =   "Text13"
            Top             =   8160
            Width           =   5295
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   2160
            TabIndex        =   20
            Text            =   "Text13"
            Top             =   8880
            Width           =   5295
         End
         Begin VB.TextBox Text13 
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
            Index           =   4
            Left            =   6000
            TabIndex        =   19
            Text            =   "Text13"
            Top             =   5880
            Width           =   1455
         End
         Begin VB.TextBox Text13 
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
            Index           =   5
            Left            =   2160
            TabIndex        =   18
            Text            =   "Text13"
            Top             =   5880
            Width           =   3735
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Index           =   0
            Left            =   10320
            TabIndex        =   17
            Text            =   "Text15"
            Top             =   1560
            Width           =   5295
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   1
            Left            =   10320
            TabIndex        =   16
            Text            =   "Text15"
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   2
            Left            =   10320
            TabIndex        =   15
            Text            =   "Text15"
            Top             =   3120
            Width           =   5295
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   3
            Left            =   10320
            TabIndex        =   14
            Text            =   "Text15"
            Top             =   3840
            Width           =   5295
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   4
            Left            =   14520
            TabIndex        =   13
            Text            =   "Text15"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   5
            Left            =   10320
            TabIndex        =   12
            Text            =   "Text15"
            Top             =   840
            Width           =   4095
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   855
            Index           =   0
            Left            =   10320
            TabIndex        =   11
            Text            =   "Text16"
            Top             =   6600
            Width           =   5295
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   1
            Left            =   10320
            TabIndex        =   10
            Text            =   "Text16"
            Top             =   7560
            Width           =   1215
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Index           =   2
            Left            =   10320
            TabIndex        =   9
            Text            =   "Text16"
            Top             =   8160
            Width           =   5295
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Index           =   3
            Left            =   10320
            TabIndex        =   8
            Text            =   "Text16"
            Top             =   8880
            Width           =   5295
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Index           =   4
            Left            =   14520
            TabIndex        =   7
            Text            =   "Text16"
            Top             =   5880
            Width           =   1095
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Index           =   5
            Left            =   10320
            TabIndex        =   6
            Text            =   "Text16"
            Top             =   5880
            Width           =   4095
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Index           =   6
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox Text13 
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
            Index           =   6
            Left            =   4200
            TabIndex        =   4
            Text            =   "Text13"
            Top             =   5280
            Width           =   3255
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Index           =   6
            Left            =   12480
            TabIndex        =   3
            Text            =   "Text15"
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "����"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Index           =   6
            Left            =   12480
            TabIndex        =   2
            Text            =   "Text16"
            Top             =   5280
            Width           =   3135
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   960
            TabIndex        =   53
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "ʵ�ʳ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   960
            TabIndex        =   52
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��Ҫ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   960
            TabIndex        =   51
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "�������"
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
            Index           =   1
            Left            =   960
            TabIndex        =   50
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   960
            TabIndex        =   49
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackColor       =   &H000000FF&
            Caption         =   "  ��̨���1"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   960
            TabIndex        =   48
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label13 
            BackColor       =   &H00808080&
            Caption         =   "  ��̨���2"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   960
            TabIndex        =   47
            Top             =   5280
            Width           =   3255
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FF8080&
            Caption         =   "  ��̨���3"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   9120
            TabIndex        =   46
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label13 
            BackColor       =   &H00008000&
            Caption         =   "  ��̨���4"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   9120
            TabIndex        =   45
            Top             =   5280
            Width           =   3375
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   960
            TabIndex        =   44
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "ʵ�ʳ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   960
            TabIndex        =   43
            Top             =   8880
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��Ҫ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   7
            Left            =   960
            TabIndex        =   42
            Top             =   8160
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "�������"
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
            Index           =   8
            Left            =   960
            TabIndex        =   41
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   9
            Left            =   960
            TabIndex        =   40
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   10
            Left            =   9120
            TabIndex        =   39
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "ʵ�ʳ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   11
            Left            =   9120
            TabIndex        =   38
            Top             =   3840
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��Ҫ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   12
            Left            =   9120
            TabIndex        =   37
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "�������"
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
            Index           =   13
            Left            =   9120
            TabIndex        =   36
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   14
            Left            =   9120
            TabIndex        =   35
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   15
            Left            =   9120
            TabIndex        =   34
            Top             =   5880
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "ʵ�ʳ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   16
            Left            =   9120
            TabIndex        =   33
            Top             =   8880
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "��Ҫ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   17
            Left            =   9120
            TabIndex        =   32
            Top             =   8160
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "�������"
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
            Index           =   18
            Left            =   9120
            TabIndex        =   31
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H0000C0C0&
            Caption         =   "������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   19
            Left            =   9120
            TabIndex        =   30
            Top             =   6600
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "Formr442"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim a As String
Dim flag1 As Integer
Dim flag2 As Boolean
Dim flag3 As Boolean     ''''''''Ⱦ���жϱ���
Dim i
Dim ksjs As Integer      '''''�����ȶ�����
Dim qpys  As Integer    '''''ȥƤ��ʱ
'''''''''''''''''             PLC ����
Dim YMSCT As String 'λԪ������ѡ���־
Dim Adree As String ' Ԫ����ַ
Dim Order As Integer 'ͨѶ˳��
Dim RWorder As Integer ' ��дͨѶ˳��
Dim RWcomm As Boolean '��ȡ����
Dim ysbc As Integer '''''''�Ĵ�����ʱ����
Dim SJPD As Integer
Dim dqdz As Integer ''''''''�ж��Ƿ�����

Dim dczw1, dczw2, dczw3, dczw4, dczw5, dczw6 As Integer ''''''''�ж��Ƿ��г�������
Dim bcbl1, bcbl2, bcbl3 As Integer ''''''''���ݱ���
Dim xrld, xrld1, xrld2, xrld3 As Integer ''''''''д���ϵ���Ϣ
Dim ytsz(7) As String ''''''''''''''''''''Һ���������
Dim ztdq1(4) As String ''''''''''''''''''''1�ųƳ���״̬��ȡ����
Dim ztdq2(4) As String ''''''''''''''''''''2�ųƳ���״̬��ȡ����
Dim ztdq3(4) As String ''''''''''''''''''''�������ݱ���ɹ�����
Dim ztdq4(4) As String ''''''''''''''''''''���Ϲ����쳣
Dim ztdq5(4) As String ''''''''''''''''''''m103--m106��״ֵ̬
Dim ctbh As String    ''''''''''''''''''''��̨���
Dim czbc As String   '''''''''''''''''''''������������
Dim zjmd As Single    '''''''''''''''''''''''�����ܶ�
Dim wdbl As String   '''''''''''''''''''''������ȡ���ȶ�����
Dim ssxscsData(35) As Single   ''''''ʵʱ��ʾ����1
Dim csfh  As Integer ''''''''''''''''''''����Һλ���
Dim ssxsData(35) As Single
Dim csfhdz(35)  As Integer   ''''''''''''''''''''����Һλ�Ĵ���
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   '����������
Dim MXH  As Integer    '''''''''ѭ����M

Private Sub Command1_Click()    '''Ԫ����λ
  Adree = YMSCT & Text4.Text
  a = gk528SetDevice(Adree, 0)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 8
  RWcomm = True
End Sub

Private Sub Command10_Click()
On Error Resume Next
  Dim b As String
  Dim COM1 As Integer
  
  COM1 = Combo1.ListIndex + 1
  b = OpenComm(MSComm4, COM1, "9600,e,7,1")
  If b = 0 Then
     Order = 0
     Timer3.Enabled = True
     Timer4.Enabled = True
     RWcomm = False
 Else
     Timer4.Enabled = False
     Timer3.Enabled = False
End If

End Sub

Private Sub Command11_Click()
On Error Resume Next
 Dim b As String
 b = CloseComm(MSComm4)
 Timer3.Enabled = False
 Timer4.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Command4_Click()
On Error Resume Next
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct isnull(��̨,'') as ��̨ FROM v_pldr_yt WHERE Ⱦ������ like '%����%' and (�������='N' or ������� is null) AND cast(CONVERT(varchar(120),��������,23) as datetime)  between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and Ⱦ�������� in(select Һ������ from ytsb) ORDER BY ��̨"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct isnull(��̨,'') as ��̨ FROM v_pldr_yt WHERE Ⱦ������ like '%����%' and �������='Y' AND cast(CONVERT(varchar(120),��������,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and Ⱦ�������� in(select Һ������ from ytsb) ORDER BY ��̨"
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



Private Sub Command6_Click()  ''''Ԫ����λ
  Adree = YMSCT & Text4.Text
  a = gk528SetDevice(Adree, 1)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True

End Sub

Private Sub Command7_Click()    '''��ѯԪ��״̬
 Adree = YMSCT & Text4.Text
 a = gk528ReadDevice(Adree, 1)  '��ַ  ����
 RWorder = 9
 RWcomm = True
End Sub

Private Sub Command8_Click()       ''''''��Ԫ��
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
    a = gk528ReadDevice(Adree, 1)  '��ַ  ����
 Else
    a = gk528ReadDevice(Adree, 2)
 End If
 RWorder = 5
 RWcomm = True
End Sub

Private Sub Command9_Click()   '''''' дԪ��
 Dim Number As String
    'д����ֵ
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
 
 If Option4.value = True Then '16λ
    ReDim WriteData(0) As String
    WriteData(0) = Val(Text7.Text)
    a = gk528WriteDevice(Adree, 1, WriteData)   '��ַ  ����  ��ֵ��
 End If
 RWorder = 6
 RWcomm = True
End Sub



Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date
DTPicker2.value = Date

'Label4.Caption = ""
If yhm = "C" Or yhm = "c" Then
Picture2(0).Visible = False
Text3.Enabled = False
Label2(3).Enabled = False
End If

For i = 0 To 5
Option15(i).Visible = True
Next
Text17 = ""
Text3 = ""
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Set RD = New ADODB.Recordset

For m = 0 To 6
Text1(m) = ""
Text13(m) = ""
Text15(m) = ""
Text16(m) = ""
Next
csfh = 1     '''''''''''''���䷢��  Һλ���
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text12 = ""
Text14 = ""
For i = 0 To 2
Text11(i) = ""
Next

For i = 0 To 9
Label12(i).Visible = False
Next
For i = 0 To 29
Label11(i).Visible = False
Next

Option16.value = True

  Dim b As String
  
  b = OpenComm(MSComm4, 4, "9600,e,7,1")
  
  If b = 0 Then
'     Label2(4).Caption = "�����Ѵ�"
     Order = 0
     Timer3.Enabled = True
     Timer4.Enabled = True
     RWcomm = False
 Else
     Label2(4).Caption = "���ڹر�"
     Timer4.Enabled = False
     Timer3.Enabled = False
 End If


    Dim g As Integer
      '*���ͨѶ��ѡ�����
      
    For g = 1 To 10                             '*���ͨѶ��ѡ��
        Combo1.AddItem "Com" & CStr(g)
    Next g
    Combo1.ListIndex = 0  '��ʾ��һ��
    Option7.value = True
    YMSCT = "M"
    DCT = "D"



Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct ����,����,�ϵ����,��������,������� FROM pldr WHERE Ⱦ������ like '%����%' and (�������='N' or ������� is null) AND �������� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY �ϵ����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct ����,����,�ϵ����,��������,������� FROM pldr WHERE Ⱦ������ like '%����%' and �������='Y' AND �������� between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY �ϵ����"
Adodc1.Refresh
End If

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"


Text2.TabIndex = 0
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(2) = 5500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 2000
VSFlexGrid2.ColWidth(5) = 2000
VSFlexGrid2.ColWidth(6) = 1000
VSFlexGrid2.ColWidth(7) = 1000

VSFlexGrid2.RowHeightMin = 600
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(4) = 2500

End Sub



Private Sub Label10_Click()
Text3 = ""
Text9 = ""
Text12 = ""
Text17 = ""
Text2.SetFocus
End Sub


Private Sub Label11_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct �ϵ����,���� FROM v_pldr_yt WHERE Ⱦ������ like '%����%' and (�������='N' or ������� is null) AND cast(CONVERT(varchar(120),��������,23) as datetime)  between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and Ⱦ�������� in(select Һ������ from ytsb) and isnull(��̨,'')='" & Label11(Index).Caption & "'  ORDER BY �ϵ���� desc"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct �ϵ����,���� FROM v_pldr_yt WHERE Ⱦ������ like '%����%' and �������='Y' AND cast(CONVERT(varchar(120),��������,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and Ⱦ�������� in(select Һ������ from ytsb) and isnull(��̨,'')='" & Label11(Index).Caption & "'  ORDER BY �ϵ���� desc"
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

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 3
pmbl = 6
Formr440.Text1 = Text3
Formr440.Show
End Select
End Sub




Private Sub MSComm4_OnComm()
 On Error Resume Next
 Dim b As String
 Dim i As Integer
 Dim Tdata1 As String, Tdata2 As String, Tdata3 As String, Tdata4 As String '*��ʱ����
 Dim Ddata(6) As Long '*�м����
 Dim Mdata(1) As Integer '*�м����
                      Dim Data10 As Long    '*�������м䴦�������
                      Dim Data As Single    '*�������м䴦�������
                      Dim dataCl As String  '*�������м䴦�������
    
 
   b = MSCONComm(MSComm4)
   Text6.Text = b
   If b <> "0" Then Exit Sub
   Timer4.Enabled = False
   Select Case Order
          Case 0   'read d68-75   ״̬��ȡ
          
                         'For i = 0 To 3
                         'dataCl = Mid(PLCText, i * 8 + 7, 2) + Mid(PLCText, i * 8 + 5, 2) + Mid(PLCText, i * 8 + 3, 2) + Mid(PLCText, i * 8 + 1, 2)
                         'Data10 = Val("&H" & dataCl) '*ת��Ϊʮ����
                         'CopyMemory Data, Data10, 4 '*ת��Ϊ������������ģ���е�COPY����,����Ϊ:��L���Ƹ�F1,λ��Ϊ��λ.
                         'ztdq1(i + 1) = Format(CStr(Data), "#0.000")
                         'Next i
                         
                         For i = 0 To 7
                         Ddata(i) = "&H" + Mid(PLCText, i * 4 + 3, 2) + Mid(PLCText, i * 4 + 1, 2)
                            '*PLC���صļĴ�����ֵ�Ǵӵ��ֽڵ����ֽ����У���������Ҫ��������һ�£�
                         If i = 0 Then
                         ztdq1(1) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 1 Then
                         ztdq4(1) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 2 Then
                         ztdq1(2) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 3 Then
                         ztdq4(2) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 4 Then
                         ztdq1(3) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 5 Then
                         ztdq4(3) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 6 Then
                         ztdq1(4) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         If i = 7 Then
                         ztdq4(4) = Format(CStr(Val(Ddata(i))), "#0")
                         End If
                         Next i
                         
                         If ztdq1(1) = 1 Then
                         Label8(1) = "���������С�������"
                         End If
                         If ztdq1(1) = 2 Then
                         Label8(1) = "����������ɡ�����"
                         End If
                        
                         If ztdq1(2) = 1 Then
                         Label8(2) = "���������С�������"
                         End If
                         If ztdq1(2) = 2 Then
                         Label8(2) = "����������ɡ�����"
                         End If
                         
                         If ztdq1(3) = 1 Then
                         Label8(3) = "���������С�������"
                         End If
                         If ztdq1(3) = 2 Then
                         Label8(3) = "����������ɡ�����"
                         End If
                         
                         If ztdq1(4) = 1 Then
                         Label8(4) = "���������С�������"
                         End If
                         If ztdq1(4) = 2 Then
                         Label8(4) = "����������ɡ�����"
                         End If
                         
                         
                         
                         
                         Order = 1
                        'read d60--d67
          Case 1
                         For i = 0 To 3
                         dataCl = Mid(PLCText, i * 8 + 7, 2) + Mid(PLCText, i * 8 + 5, 2) + Mid(PLCText, i * 8 + 3, 2) + Mid(PLCText, i * 8 + 1, 2)
                         Data10 = Val("&H" & dataCl) '*ת��Ϊʮ����
                        ' CopyMemory Data, Data10, 4 '*ת��Ϊ������������ģ���е�COPY����,����Ϊ:��L���Ƹ�F1,λ��Ϊ��λ.
                         ztdq2(i + 1) = Format(CStr(Data10) / 1000, "#0.000")
                         Next i
                         
                         Text1(3) = Format(Val(ztdq2(1)), "#0.000")
                         Text13(3) = Format(Val(ztdq2(2)), "#0.000")
                         Text15(3) = Format(Val(ztdq2(3)), "#0.000")
                         Text16(3) = Format(Val(ztdq2(4)), "#0.000")
                         
                         
                         Order = 0
          Case 2
                        'Tdata1 = Mid(PLCText, 1, 2)
                        'Mdata(0) = Val("&H" + Tdata1) ' ת��Ϊʮ����
                        'Tdata2 = DecToBin(Mdata(0)) '*���ö�����ת������
                        'Tdata3 = StrReverse(Tdata2)
                      
                        'Tdata1 = Mid(PLCText, 3, 2)
                        'Mdata(0) = Val("&H" + Tdata1) ' ת��Ϊʮ����
                        'Tdata2 = DecToBin(Mdata(0)) '*���ö�����ת������
                        'Tdata4 = StrReverse(Tdata2)
                      
                        'Tdata2 = Tdata3 + Tdata4
                        'For i = 0 To 3
                        'If Mid(Tdata2, i + 1, 1) = 1 Then
                        'ztdq5(i + 1) = "1"            '''��m103---m106 ��ֵ
                        'Else
                        'ztdq5(i + 1) = "0"     ' &H80FF80
                        'End If
                        'Next i
                        'Order = 0
                         
         Case 6, 7, 8  'д �ã���λ
               Order = 0
   End Select

   Timer3.Enabled = True

End Sub


Private Sub Option15_Click(Index As Integer)
Select Case Index
       Case Index
If Option15(Index).value = True Then
If Text9 <> Option15(Index).Caption Then
Text9 = Option15(Index).Caption
Text1(0) = ""
Text1(1) = ""
End If
End If
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       
       Case 4
If Text1(4) = "0" Then
Timer1.Enabled = False
End If

If Text1(4) = "1" Then
Beep 2000, 50
qpys = 3
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
    a = gk528ReadDevice(Adree, 1)  '��ַ  ����
 Else
    a = gk528ReadDevice(Adree, 2)
 End If
 RWorder = 5
 RWcomm = True
End Sub


Private Sub Text11_Change(Index As Integer)
Select Case Index
       Case 0
If Val(Text1(3)) > 0 And Val(Text11(0)) = 2 Then
Timer2.Enabled = True
ksjs = 0
End If
       Case 2
    If Val(Text11(2)) = 1 Then
    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "clyclr('" & Text3 & "','" & Text9 & "','" & Text12 & "','" & Text1(1) & "','" & Now & "')"    ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
    End If
Text11(2) = ""
End Select
End Sub

Private Sub Text12_Change()
'If Val(ztdq(0)) <> 0 Then
'Text3 = ""
'Label4.Caption = "��������û����ɣ���ֹɨ���ϵ���"
'Exit Sub
'End If
'Call SGJC
End Sub



Private Sub Text13_Change(Index As Integer)
Select Case Index
       
       Case 4
If Text13(4) = "0" Then
Timer1.Enabled = False
End If

If Text13(4) = "1" Then
Beep 2000, 50
qpys = 1
Timer1.Enabled = True
Text13(4) = ""
End If
End Select

End Sub

Private Sub Text15_Change(Index As Integer)
Select Case Index
       
       Case 4
If Text15(4) = "0" Then
Timer1.Enabled = False
End If

If Text15(4) = "1" Then
Beep 2000, 50
qpys = 1
Timer1.Enabled = True
Text15(4) = ""
End If
End Select

End Sub

Private Sub Text16_Change(Index As Integer)
Select Case Index
       Case 4
If Text16(4) = "0" Then
Timer1.Enabled = False
End If

If Text16(4) = "1" Then
Beep 2000, 50
qpys = 1
Timer1.Enabled = True
Text16(4) = ""
End If
End Select

End Sub

Private Sub Text17_Change()
Call SGJC
End Sub

Private Sub Text2_Change()
If Len(Text2) = 10 Then
Adodc7.RecordSource = "select �ϵ���� from ldkh where ���ű��='" & Text2 & "'"
Adodc7.Refresh
If Adodc7.Recordset.EOF Then
Text2 = ""
Text2.SetFocus
Else
Text3 = Adodc7.Recordset.Fields(0)
Text2 = ""
Text2.SetFocus
End If
End If

If InStr(Text2, "J") > 0 Then
Text3 = Mid(Text2, 1, Len(Text2) - 1)
Text2 = ""
Text2.SetFocus
End If
End Sub

Private Sub Text3_Change()
'On Error Resume Next

For i = 0 To 5
Option15(i).value = False
Option15(i).Visible = False
Next
Text9 = ""
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Adodc4.RecordSource = "SELECT distinct �������� FROM pldr WHERE �ϵ����='" & Text3 & "' and Ⱦ������ like '%����%' ORDER BY ��������"
Adodc4.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Adodc4.Recordset.EOF Then
For i = 0 To 5
Option15(i).Visible = False
Next
Else
i = 0
Do While Not Adodc4.Recordset.EOF
Option15(i).Caption = Adodc4.Recordset.Fields(0)
Option15(i).Visible = True
i = i + 1
Adodc4.Recordset.MoveNext
Loop
End If

End Sub

Private Sub Text4_Change()
  Adree = "M66"
  a = gk528SetDevice(Adree, 1)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
End Sub



Private Sub Text7_Change()      ''''''''''''''��ʱ����  �� timer
'If Val(Text7) > 100 Then
'    Adree = "D" & Text5.Text
'    ReDim WriteData(0) As String
'    WriteData(0) = Val(Text7.Text)
'    A = gk528WriteDevice(Adree, 1, WriteData)   '��ַ  ����  ��ֵ��
' RWorder = 6
' RWcomm = True
'Text1(4) = "1"
'End If
End Sub

Private Sub Text8_Change()
Text1(4) = Text8
End Sub


Private Sub Text9_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Adodc2.RecordSource = "SELECT �������,Ⱦ��������,���ϵ�λ,��������,ʵ�ʳ���,�����,��̨�� as �ӱ�� FROM v_pldr_yt WHERE �ϵ����='" & Text3 & "' and Ⱦ������ like '%����%' and ��������='" & Text9 & "' ORDER BY ��������,�����"
Adodc2.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(2) = 5500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 2000
VSFlexGrid2.ColWidth(5) = 2000
VSFlexGrid2.ColWidth(6) = 1000
VSFlexGrid2.ColWidth(7) = 1000
'Call VQJC
End Sub

Private Sub Timer1_Timer()
       Dim Number As String
       Dim WriteData() As String
       ReDim WriteData(0 To 14) As String  ''''''д�����
'       Dim DataW As String    '*���������м䴦�������
'       Dim Data10(7) As Single   '*���������м䴦�������
'       Dim Buffer(3) As Byte   '*���������м䴦�������

If ctbh = "1" Then    ''''''''''''''''''''''''''''''''''''''''''''''''''''''1�ų�
 If qpys = 1 Then    ''''''����׼������ʱ
           
           
'       For i = 0 To 1
'       Data10(i) = Val(ytsz(i))
'       CopyMemory Buffer(0), Data10(i), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
'       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
'       Next
'       a = gk528WriteDevice("D480", 4, WriteData())
       For i = 0 To 1
       Number = Right("00000000" + Hex(Val(ytsz(i))), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D100", 4, WriteData())
 RWorder = 6
 RWcomm = True
 Timer1.Enabled = False
 Timer5.Enabled = True
 End If


qpys = qpys - 1
Label8(1).Caption = "��ע�������׼���У���" + Trim(qpys)
End If


If ctbh = "2" Then    ''''''''''''''''''''''''''''''''''''''''''''''''''''2�ų�
If qpys = 1 Then    ''''''����׼������ʱ
           
           
'       For i = 0 To 1
'       Data10(i) = Val(ytsz(i))
'       CopyMemory Buffer(0), Data10(i), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
'       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
'       Next
'       a = gk528WriteDevice("D484", 4, WriteData())
       For i = 0 To 1
       Number = Right("00000000" + Hex(Val(ytsz(i))), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D104", 4, WriteData())
 RWorder = 6
 RWcomm = True
 
 Timer1.Enabled = False
 Timer5.Enabled = True
 End If


qpys = qpys - 1
Label8(2).Caption = "��ע�������׼���У���" + Trim(qpys)
End If

If ctbh = "3" Then   '''''''''''''''''''''''''''''''''''''''''''''''''''''3�ų�
If qpys = 1 Then    ''''''����׼������ʱ
           
           
'       For i = 0 To 1
'       Data10(i) = Val(ytsz(i))
'       CopyMemory Buffer(0), Data10(i), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
'       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
'       Next
'       a = gk528WriteDevice("D488", 4, WriteData())
       For i = 0 To 1
       Number = Right("00000000" + Hex(Val(ytsz(i))), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D108", 4, WriteData())
 RWorder = 6
 RWcomm = True
 
 Timer1.Enabled = False
 Timer5.Enabled = True
 End If


qpys = qpys - 1
Label8(3).Caption = "��ע�������׼���У���" + Trim(qpys)
End If

If ctbh = "4" Then   ''''''''''''''''''''''''''''''''''''''''''''''''''''4�ų�
If qpys = 1 Then    ''''''����׼������ʱ
           
           
'       For i = 0 To 1
'       Data10(i) = Val(ytsz(i))
'       CopyMemory Buffer(0), Data10(i), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
'       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
'       Next
'       a = gk528WriteDevice("D492", 4, WriteData())
       For i = 0 To 1
       Number = Right("00000000" + Hex(Val(ytsz(i))), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D112", 4, WriteData())
 RWorder = 6
 RWcomm = True
 
 Timer1.Enabled = False
 Timer5.Enabled = True
 End If


qpys = qpys - 1
Label8(4).Caption = "��ע�������׼���У���" + Trim(qpys)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Timer2_Timer()

If Val(ztdq1(1)) = 2 And Val(Text1(3)) > 0 Then
cll = Format(Val(Text1(3)), "#0.000")   ''''''''''������λgת����kg
sql1 = "UPDATE pldr SET ʵ�ʳ���='" & cll & "',�������='Y',��������='" & Now & "' WHERE �ϵ����='" & Text1(6) & "' and Ⱦ��������='" & Text1(0) & "' and �����='" & Text1(1) & "' and ��������='" & Text1(5) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.Refresh
wdbl = "0"
czbc = "0"
ztdq3(1) = "1"
Timer5.Enabled = False
Timer6.Enabled = True
End If

If Val(ztdq1(2)) = 2 And Val(Text13(3)) > 0 Then
cll = Format(Val(Text13(3)), "#0.000")   ''''''''''������λgת����kg
sql1 = "UPDATE pldr SET ʵ�ʳ���='" & cll & "',�������='Y',��������='" & Now & "' WHERE �ϵ����='" & Text13(6) & "' and Ⱦ��������='" & Text13(0) & "' and �����='" & Text13(1) & "' and ��������='" & Text13(5) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.Refresh
wdbl = "0"
czbc = "0"
ztdq3(2) = "1"
Timer5.Enabled = False
Timer6.Enabled = True
End If

If Val(ztdq1(3)) = 2 And Val(Text15(3)) > 0 Then
cll = Format(Val(Text15(3)), "#0.000")   ''''''''''������λgת����kg
sql1 = "UPDATE pldr SET ʵ�ʳ���='" & cll & "',�������='Y',��������='" & Now & "' WHERE �ϵ����='" & Text15(6) & "' and Ⱦ��������='" & Text15(0) & "' and �����='" & Text15(1) & "' and ��������='" & Text15(5) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.Refresh
wdbl = "0"
czbc = "0"
ztdq3(3) = "1"
Timer5.Enabled = False
Timer6.Enabled = True
End If

If Val(ztdq1(4)) = 2 And Val(Text16(3)) > 0 Then
cll = Format(Val(Text16(3)), "#0.000")   ''''''''''������λgת����kg
sql1 = "UPDATE pldr SET ʵ�ʳ���='" & cll & "',�������='Y',��������='" & Now & "' WHERE �ϵ����='" & Text16(6) & "' and Ⱦ��������='" & Text16(0) & "' and �����='" & Text16(1) & "' and ��������='" & Text16(5) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.Refresh
wdbl = "0"
czbc = "0"
ztdq3(4) = "1"
Timer5.Enabled = False
Timer6.Enabled = True
End If

VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(2) = 5500
VSFlexGrid2.ColWidth(3) = 1500
VSFlexGrid2.ColWidth(4) = 2000
VSFlexGrid2.ColWidth(5) = 2000
VSFlexGrid2.ColWidth(6) = 1000
VSFlexGrid2.ColWidth(7) = 1000

End Sub


Private Sub Timer3_Timer()    ''''''''''''''PLC

 If RWcomm = True Then
   Order = RWorder
   RWcomm = False
 End If
  Select Case Order
         Case 0   '��D56
              a = gk528ReadDevice("D68", 8)
         Case 1
              a = gk528ReadDevice("D60", 8)
  End Select
        

 MSComm4.OutBufferCount = 0 '*���ò����ط��ͻ��������ֽ���,��Ϊ0ʱ��շ��ͻ�����
 MSComm4.InBufferCount = 0  '*���ò����ؽ��ջ��������ֽ���,��Ϊ0ʱ��ս��ջ�����
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

Private Sub Timer6_Timer()
        Dim Number As String
        'д����ֵ
        Dim WriteData() As String
       ReDim WriteData(0 To 14) As String  ''''''д�����
       Dim DataW As String    '*���������м䴦�������
       Dim Data10(7) As Single   '*���������м䴦�������
       Dim Buffer(3) As Byte   '*���������м䴦�������


If ztdq3(1) = "1" Then
''''''''''''''''''''''''''''''''''''''''''''''''���״̬
If Val(ztdq1(1)) = 2 Then  ''''And ztdq5(1) = "1" Then
'       Data10(0) = 0
'       CopyMemory Buffer(0), Data10(0), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(0) = Val("&H" + Right(DataW, 4))
'       WriteData(1) = Val("&H" + Mid(DataW, 1, 4))
'       a = gk528WriteDevice("D700", 2, WriteData())
'       RWorder = 6
'       RWcomm = True
       For i = 0 To 0
       Number = Right("00000000" + Hex(Val(0)), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D700", 2, WriteData())
       RWorder = 6
       RWcomm = True
End If

If Val(ztdq1(1)) = 0 Then  ''And ztdq5(1) = "1" Then
For i = 0 To 6
Text1(i) = ""
Next
ztdq3(1) = "0"
Label8(1) = "����׼���С�������"
End If

''If Val(ztdq1(1)) = 2 And ztdq5(1) = "0" Then
'Adodc5.RecordSource = "select * from pldr  WHERE �������='Y' and �ϵ����='" & Text1(6) & "' and Ⱦ��������='" & Text1(0) & "' and �����='" & Text1(1) & "' and ��������='" & Text1(5) & "'"
'Adodc5.Refresh
'If Not Adodc5.Recordset.EOF Then
'  A = gk528SetDevice("M103", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
'End If
'End If

End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If ztdq3(2) = "1" Then
''''''''''''''''''''''''''''''''''''''''''''''''���״̬
If Val(ztdq1(2)) = 2 Then   ''And ztdq5(2) = "1" Then
'       Data10(0) = 0
'       CopyMemory Buffer(0), Data10(0), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(0) = Val("&H" + Right(DataW, 4))
'       WriteData(1) = Val("&H" + Mid(DataW, 1, 4))
'       a = gk528WriteDevice("D702", 2, WriteData())
       For i = 0 To 0
       Number = Right("00000000" + Hex(Val(0)), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D702", 2, WriteData())
       RWorder = 6
       RWcomm = True
End If
'If Val(ztdq1(2)) = 2 And ztdq5(2) = "0" Then
'Adodc5.RecordSource = "select * from pldr  WHERE �������='Y' and �ϵ����='" & Text13(6) & "' and Ⱦ��������='" & Text13(0) & "' and �����='" & Text13(1) & "' and ��������='" & Text13(5) & "'"
'Adodc5.Refresh
'If Not Adodc5.Recordset.EOF Then
'  A = gk528SetDevice("M104", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
'End If
'End If

If Val(ztdq1(2)) = 0 Then  ''And ztdq5(2) = "1" Then
For i = 0 To 6
Text13(i) = ""
Next
ztdq3(2) = "0"
Label8(2) = "����׼���С�������"
End If

End If

'''''''''''''''''''''''''''''''''''''''''''''''''
If ztdq3(3) = "1" Then
''''''''''''''''''''''''''''''''''''''''''''''''���״̬
If Val(ztdq1(3)) = 2 Then    '''And ztdq5(3) = "1" Then
'       Data10(0) = 0
'       CopyMemory Buffer(0), Data10(0), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(0) = Val("&H" + Right(DataW, 4))
'       WriteData(1) = Val("&H" + Mid(DataW, 1, 4))
'       a = gk528WriteDevice("D704", 2, WriteData())
       For i = 0 To 0
       Number = Right("00000000" + Hex(Val(0)), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D704", 2, WriteData())
       RWorder = 6
       RWcomm = True
End If
'If Val(ztdq1(3)) = 2 And ztdq5(3) = "1" Then
'Adodc5.RecordSource = "select * from pldr  WHERE �������='Y' and �ϵ����='" & Text15(6) & "' and Ⱦ��������='" & Text15(0) & "' and �����='" & Text15(1) & "' and ��������='" & Text15(5) & "'"
'Adodc5.Refresh
'If Not Adodc5.Recordset.EOF Then
'  A = gk528SetDevice("M105", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
'End If
'End If

If Val(ztdq1(3)) = 0 Then   ''''And ztdq5(3) = "1" Then
For i = 0 To 6
Text15(i) = ""
Next
ztdq3(3) = "0"
Label8(3) = "����׼���С�������"
End If

End If

''''''''''''''''''''''''''''''''''''''''''''''''
If ztdq3(4) = "1" Then
''''''''''''''''''''''''''''''''''''''''''''''''���״̬
If Val(ztdq1(4)) = 2 Then   '''And ztdq5(4) = "1" Then
'       Data10(0) = 0
'       CopyMemory Buffer(0), Data10(0), 4
'       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
'       WriteData(0) = Val("&H" + Right(DataW, 4))
'       WriteData(1) = Val("&H" + Mid(DataW, 1, 4))
'       a = gk528WriteDevice("D706", 2, WriteData())
       For i = 0 To 0
       Number = Right("00000000" + Hex(Val(0)), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Number, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Number, 1, 4))
       Next
       a = gk528WriteDevice("D706", 2, WriteData())
       RWorder = 6
       RWcomm = True
End If
'If Val(ztdq1(4)) = 2 And ztdq5(4) = "1" Then
'Adodc5.RecordSource = "select * from pldr  WHERE �������='Y' and �ϵ����='" & Text16(6) & "' and Ⱦ��������='" & Text16(0) & "' and �����='" & Text16(1) & "' and ��������='" & Text16(5) & "'"
'Adodc5.Refresh
'If Not Adodc5.Recordset.EOF Then
'  A = gk528SetDevice("M106", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
'End If
'End If

If Val(ztdq1(4)) = 0 Then   ''And ztdq5(4) = "1" Then
For i = 0 To 6
Text16(i) = ""
Next
ztdq3(4) = "0"
Label8(4) = "����׼���С�������"
End If

End If
End Sub

Private Sub VSFlexGrid1_dblClick()
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
Text3 = Adodc1.Recordset.Fields(2)
End Sub

Private Sub SGJC()
'On Error Resume Next
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=hsfw"
Adodc3.RecordSource = "SELECT ISNULL(�������,'N'),��������,Ⱦ������,Ⱦ��������,���ϵ�λ,round(��������,6),ʵ�ʳ���,�����,��̨��,��ϴˮ��,Һ���ܶ�,����ʱ��,�ܵ����,Һ���� FROM v_pldr_yt WHERE (�������<>'Y' OR ������� IS NULL) and  �ϵ����='" & Text3 & "' and Ⱦ������ like '%����%' and ��������='" & Text9 & "' AND Ⱦ��������='" & Text12 & "' AND �����='" & Text17 & "' ORDER BY ��������,�����"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
'Text1(0) = ""
'Text1(1) = ""
'Text1(2) = ""
'Text1(3) = ""
'Text1(4) = ""
''''''''''''''''''''''''Label4.Caption = "�������"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��������λ
'Text4 = ""
For i = 0 To 6
ytsz(i) = ""
Next
wdbl = "0"
Else
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ݳ���ת������
If Adodc3.Recordset.Fields(0) <> "Y" And Adodc3.Recordset.Fields(8) = "1" And Val(ztdq1(1)) = 0 Then
Text1(0) = Adodc3.Recordset.Fields(3)
Text1(1) = Adodc3.Recordset.Fields(7)
Text1(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)), "#0.000")
Text1(3) = 0
Text1(5) = Adodc3.Recordset.Fields(1)
Text1(6) = Text3
If Option16.value = True Then
ytsz(0) = Adodc3.Recordset.Fields(12)
Else
ytsz(0) = Adodc3.Recordset.Fields(12)     ''''''�ܵ����
End If
ytsz(1) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000, "#0")  ''''''������
ctbh = Adodc3.Recordset.Fields(8)
wdbl = "1"                                '''''''''�ȶ�����
Text1(4) = "1"                            ''''''''''''''''д����
Exit Sub
End If

If Adodc3.Recordset.Fields(0) <> "Y" And Adodc3.Recordset.Fields(8) = "2" And Val(ztdq1(2)) = 0 Then
Text13(0) = Adodc3.Recordset.Fields(3)
Text13(1) = Adodc3.Recordset.Fields(7)
Text13(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)), "#0.000")
Text13(3) = 0
Text13(5) = Adodc3.Recordset.Fields(1)
Text13(6) = Text3
If Option16.value = True Then
ytsz(0) = Adodc3.Recordset.Fields(12)
Else
ytsz(0) = Adodc3.Recordset.Fields(12)
End If
ytsz(1) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000, "#0")  ''''''������
ctbh = Adodc3.Recordset.Fields(8)

wdbl = "1"                                '''''''''�ȶ�����
Text13(4) = "1"                            ''''''''''''''''д����
Exit Sub
End If



If Adodc3.Recordset.Fields(0) <> "Y" And Adodc3.Recordset.Fields(8) = "3" And Val(ztdq1(3)) = 0 Then
Text15(0) = Adodc3.Recordset.Fields(3)
Text15(1) = Adodc3.Recordset.Fields(7)
Text15(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)), "#0.000")
Text15(3) = 0
Text15(5) = Adodc3.Recordset.Fields(1)
Text15(6) = Text3
If Option16.value = True Then
ytsz(0) = Adodc3.Recordset.Fields(12)
Else
ytsz(0) = Adodc3.Recordset.Fields(12)
End If
ytsz(1) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000, "#0")  ''''''������
ctbh = Adodc3.Recordset.Fields(8)

wdbl = "1"                                '''''''''�ȶ�����
Text15(4) = "1"                            ''''''''''''''''д����
Exit Sub
End If



If Adodc3.Recordset.Fields(0) <> "Y" And Adodc3.Recordset.Fields(8) = "4" And Val(ztdq1(4)) = 0 Then
Text16(0) = Adodc3.Recordset.Fields(3)
Text16(1) = Adodc3.Recordset.Fields(7)
Text16(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)), "#0.000")
Text16(3) = 0
Text16(5) = Adodc3.Recordset.Fields(1)
Text16(6) = Text3
If Option16.value = True Then
ytsz(0) = Adodc3.Recordset.Fields(12)
Else
ytsz(0) = Adodc3.Recordset.Fields(12)
End If
ytsz(1) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000, "#0")  ''''''������
ctbh = Adodc3.Recordset.Fields(8)

wdbl = "1"                                '''''''''�ȶ�����
Text16(4) = "1"                            ''''''''''''''''д����
Exit Sub
End If
Adodc3.Recordset.MoveNext
Loop
End If


End Sub

Private Sub VSFlexGrid2_Click()
On Error Resume Next
If Adodc2.Recordset.EOF Then Exit Sub
Adodc2.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc2.Recordset.Move rs - 1
Text17 = ""
Text12 = Adodc2.Recordset.Fields(1)
Text17 = Adodc2.Recordset.Fields(5)
End Sub


