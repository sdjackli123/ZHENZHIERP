VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formj1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "销售合同"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15810
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   19711
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   1058
      ForeColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "订单明细"
      TabPicture(0)   =   "Formj1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "加工信息"
      TabPicture(1)   =   "Formj1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "订单工序"
      TabPicture(2)   =   "Formj1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "订单附注"
      TabPicture(3)   =   "Formj1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "订单作废"
      TabPicture(4)   =   "Formj1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Left            =   -75000
         ScaleHeight     =   10275
         ScaleWidth      =   15555
         TabIndex        =   202
         Top             =   600
         Width           =   15615
         Begin MSAdodcLib.Adodc Adodc31 
            Height          =   375
            Left            =   4680
            Top             =   9240
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
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
            Caption         =   "Adodc31"
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
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   2640
            TabIndex        =   237
            Text            =   "Text10"
            Top             =   600
            Width           =   3375
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
            Bindings        =   "Formj1.frx":008C
            Height          =   7455
            Left            =   480
            TabIndex        =   235
            Top             =   1320
            Width           =   14535
            _cx             =   25638
            _cy             =   13150
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
         Begin VB.CommandButton Command34 
            BackColor       =   &H00C0C0FF&
            Caption         =   "单号恢复"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   234
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Height          =   495
            Index           =   1
            Left            =   12960
            TabIndex        =   208
            Text            =   "Text8"
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command28 
            BackColor       =   &H00C0C0FF&
            Caption         =   "单号删除"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13920
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   206
            Top             =   1080
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command27 
            BackColor       =   &H00C0C0FF&
            Caption         =   "单号复制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13920
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   205
            Top             =   240
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            Height          =   495
            Index           =   0
            Left            =   12960
            TabIndex        =   204
            Text            =   "Text8"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "单号"
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
            Left            =   1680
            TabIndex        =   238
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFC0&
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
            Height          =   495
            Left            =   480
            TabIndex        =   236
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C0C0&
            Caption         =   "要删除单号"
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
            Index           =   1
            Left            =   12240
            TabIndex        =   207
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C0C0&
            Caption         =   "被复制单号"
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
            Left            =   12240
            TabIndex        =   203
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0E0FF&
         Height          =   10575
         Left            =   -75000
         ScaleHeight     =   10515
         ScaleWidth      =   15555
         TabIndex        =   104
         Top             =   600
         Visible         =   0   'False
         Width           =   15615
         Begin VB.CommandButton Command25 
            BackColor       =   &H00C0C0FF&
            Caption         =   "合同打印"
            Height          =   615
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   197
            Top             =   7920
            Width           =   1455
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00C0C0FF&
            Caption         =   "附注删除"
            Height          =   615
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   6840
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "来料规格"
            Height          =   2535
            Index           =   0
            Left            =   120
            TabIndex        =   145
            Top             =   1560
            Width           =   9975
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   0
               Left            =   1560
               TabIndex        =   172
               Text            =   "Text24"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   1
               Left            =   2880
               TabIndex        =   171
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   2
               Left            =   3720
               TabIndex        =   170
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   3
               Left            =   4560
               TabIndex        =   169
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   4
               Left            =   5400
               TabIndex        =   168
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   5
               Left            =   6240
               TabIndex        =   167
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   6
               Left            =   7080
               TabIndex        =   166
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   7
               Left            =   7920
               TabIndex        =   165
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text24 
               Height          =   375
               Index           =   8
               Left            =   8760
               TabIndex        =   164
               Text            =   "Text24"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   0
               Left            =   1560
               TabIndex        =   163
               Text            =   "Text25"
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   1
               Left            =   2880
               TabIndex        =   162
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   2
               Left            =   3720
               TabIndex        =   161
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   3
               Left            =   4560
               TabIndex        =   160
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   4
               Left            =   5400
               TabIndex        =   159
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   5
               Left            =   6240
               TabIndex        =   158
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   6
               Left            =   7080
               TabIndex        =   157
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   7
               Left            =   7920
               TabIndex        =   156
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text25 
               Height          =   375
               Index           =   8
               Left            =   8760
               TabIndex        =   155
               Text            =   "Text25"
               Top             =   1320
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   0
               Left            =   1560
               TabIndex        =   154
               Text            =   "Text26"
               Top             =   1800
               Width           =   1215
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   1
               Left            =   2880
               TabIndex        =   153
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   2
               Left            =   3720
               TabIndex        =   152
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   3
               Left            =   4560
               TabIndex        =   151
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   4
               Left            =   5400
               TabIndex        =   150
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   5
               Left            =   6240
               TabIndex        =   149
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   6
               Left            =   7080
               TabIndex        =   148
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   7
               Left            =   7920
               TabIndex        =   147
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox Text26 
               Height          =   375
               Index           =   8
               Left            =   8760
               TabIndex        =   146
               Text            =   "Text26"
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "面料"
               Height          =   375
               Index           =   20
               Left            =   240
               TabIndex        =   184
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "辅料"
               Height          =   375
               Index           =   21
               Left            =   240
               TabIndex        =   183
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "其它"
               Height          =   375
               Index           =   22
               Left            =   240
               TabIndex        =   182
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "品种"
               Height          =   375
               Index           =   23
               Left            =   1560
               TabIndex        =   181
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "纱支"
               Height          =   375
               Index           =   24
               Left            =   2880
               TabIndex        =   180
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "成分"
               Height          =   375
               Index           =   25
               Left            =   3720
               TabIndex        =   179
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "混纺比"
               Height          =   375
               Index           =   26
               Left            =   4560
               TabIndex        =   178
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "幅宽"
               Height          =   375
               Index           =   27
               Left            =   5400
               TabIndex        =   177
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "克重"
               Height          =   375
               Index           =   28
               Left            =   6240
               TabIndex        =   176
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "数量"
               Height          =   375
               Index           =   29
               Left            =   7080
               TabIndex        =   175
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "件数"
               Height          =   375
               Index           =   30
               Left            =   8760
               TabIndex        =   174
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "类型"
               Height          =   375
               Index           =   31
               Left            =   7920
               TabIndex        =   173
               Top             =   240
               Width           =   735
            End
            Begin VB.Line Line2 
               X1              =   240
               X2              =   9840
               Y1              =   720
               Y2              =   720
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "成品布面要求"
            Height          =   1455
            Index           =   3
            Left            =   120
            TabIndex        =   138
            Top             =   4320
            Width           =   6855
            Begin VB.TextBox Text12 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   141
               Text            =   "Text12"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox Text12 
               Height          =   375
               Index           =   1
               Left            =   1560
               TabIndex        =   140
               Text            =   "Text12"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox Text12 
               Height          =   375
               Index           =   2
               Left            =   3000
               TabIndex        =   139
               Text            =   "Text12"
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "级差原样"
               Height          =   375
               Index           =   33
               Left            =   120
               TabIndex        =   144
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "匹差"
               Height          =   375
               Index           =   34
               Left            =   1560
               TabIndex        =   143
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "布面疵点少于"
               Height          =   375
               Index           =   35
               Left            =   3000
               TabIndex        =   142
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "特殊要求"
            Height          =   975
            Index           =   4
            Left            =   120
            TabIndex        =   137
            Top             =   6000
            Width           =   6855
            Begin VB.TextBox Text12 
               Height          =   615
               Index           =   3
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   186
               Text            =   "Formj1.frx":00A2
               Top             =   240
               Width           =   6375
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "产品去向与要求"
            Height          =   1215
            Index           =   5
            Left            =   10200
            TabIndex        =   132
            Top             =   1680
            Width           =   4215
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   0
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   134
               Text            =   "Formj1.frx":00A9
               Top             =   240
               Width           =   2895
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   1
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   133
               Text            =   "Formj1.frx":00AF
               Top             =   720
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "内销"
               Height          =   375
               Index           =   39
               Left            =   480
               TabIndex        =   136
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "出口"
               Height          =   375
               Index           =   40
               Left            =   480
               TabIndex        =   135
               Top             =   720
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "包装方式"
            Height          =   1215
            Index           =   7
            Left            =   10200
            TabIndex        =   129
            Top             =   2880
            Width           =   4215
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   3
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   190
               Text            =   "Formj1.frx":00B5
               Top             =   720
               Width           =   2895
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   2
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   189
               Text            =   "Formj1.frx":00BB
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "内用"
               Height          =   375
               Index           =   42
               Left            =   480
               TabIndex        =   131
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "外用"
               Height          =   375
               Index           =   43
               Left            =   480
               TabIndex        =   130
               Top             =   720
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "产品交付"
            Height          =   1335
            Index           =   8
            Left            =   4800
            TabIndex        =   122
            Top             =   120
            Width           =   9615
            Begin VB.TextBox Text14 
               Height          =   375
               Index           =   0
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   125
               Text            =   "Formj1.frx":00C1
               Top             =   360
               Width           =   3855
            End
            Begin VB.TextBox Text14 
               Height          =   375
               Index           =   1
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   124
               Text            =   "Formj1.frx":00C7
               Top             =   840
               Width           =   3855
            End
            Begin VB.TextBox Text14 
               Height          =   375
               Index           =   2
               Left            =   6000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   123
               Text            =   "Formj1.frx":00CD
               Top             =   600
               Width           =   3375
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "交货地址"
               Height          =   375
               Index           =   41
               Left            =   120
               TabIndex        =   128
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "产品交提货期限"
               Height          =   375
               Index           =   44
               Left            =   120
               TabIndex        =   127
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "联系人或部门"
               Height          =   375
               Index           =   45
               Left            =   6000
               TabIndex        =   126
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "产品验收时间"
            Height          =   735
            Index           =   9
            Left            =   10200
            TabIndex        =   120
            Top             =   4320
            Width           =   4215
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   4
               Left            =   1200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   191
               Text            =   "Formj1.frx":00D3
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "验收时间"
               Height          =   375
               Index           =   46
               Left            =   480
               TabIndex        =   121
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "加工结算方式"
            Height          =   1335
            Index           =   10
            Left            =   7560
            TabIndex        =   114
            Top             =   5160
            Width           =   6855
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   8
               Left            =   5280
               ScrollBars      =   2  'Vertical
               TabIndex        =   195
               Text            =   "Text9"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   7
               Left            =   5280
               ScrollBars      =   2  'Vertical
               TabIndex        =   194
               Text            =   "Text9"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   6
               Left            =   2880
               ScrollBars      =   2  'Vertical
               TabIndex        =   193
               Text            =   "Text9"
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox Text9 
               Height          =   375
               Index           =   5
               Left            =   2880
               ScrollBars      =   2  'Vertical
               TabIndex        =   192
               Text            =   "Text9"
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "现金、账期、定金"
               Height          =   375
               Index           =   47
               Left            =   240
               TabIndex        =   119
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "结款方式"
               Height          =   375
               Index           =   48
               Left            =   2040
               TabIndex        =   118
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "金额"
               Height          =   375
               Index           =   49
               Left            =   2040
               TabIndex        =   117
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "逾期利息"
               Height          =   375
               Index           =   50
               Left            =   4440
               TabIndex        =   116
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "滞纳金"
               Height          =   375
               Index           =   51
               Left            =   4440
               TabIndex        =   115
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "责任与义务"
            Height          =   735
            Index           =   11
            Left            =   120
            TabIndex        =   112
            Top             =   7080
            Width           =   6855
            Begin VB.TextBox Text12 
               Height          =   375
               Index           =   4
               Left            =   2040
               TabIndex        =   187
               Text            =   "Text12"
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "超出领取期限"
               Height          =   375
               Index           =   52
               Left            =   240
               TabIndex        =   113
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "其它约定"
            Height          =   975
            Index           =   12
            Left            =   120
            TabIndex        =   110
            Top             =   8040
            Width           =   6855
            Begin VB.TextBox Text12 
               Height          =   375
               Index           =   5
               Left            =   2040
               TabIndex        =   188
               Text            =   "Text12"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "签订生效天数"
               Height          =   375
               Index           =   53
               Left            =   240
               TabIndex        =   111
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "签约地点"
            Height          =   1215
            Index           =   13
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   3975
            Begin VB.TextBox Text14 
               Height          =   375
               Index           =   4
               Left            =   1440
               ScrollBars      =   2  'Vertical
               TabIndex        =   198
               Text            =   "Text9"
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox Text14 
               Height          =   375
               Index           =   3
               Left            =   1440
               ScrollBars      =   2  'Vertical
               TabIndex        =   196
               Text            =   "Text9"
               Top             =   240
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   375
               Left            =   1440
               TabIndex        =   200
               Top             =   720
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarTitleBackColor=   8421376
               CalendarTrailingForeColor=   255
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   328007683
               CurrentDate     =   39177
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "日期"
               Height          =   375
               Index           =   19
               Left            =   240
               TabIndex        =   199
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               Caption         =   "地点"
               Height          =   375
               Index           =   18
               Left            =   240
               TabIndex        =   109
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H00C0C0FF&
            Caption         =   "附注保存"
            Height          =   615
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   6840
            Width           =   1455
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00C0C0FF&
            Caption         =   "附注显示"
            Height          =   615
            Left            =   9240
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   6840
            Width           =   1455
         End
         Begin VB.CommandButton Command23 
            BackColor       =   &H00C0C0FF&
            Caption         =   "附注修改"
            Height          =   615
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   6840
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0E0FF&
         Height          =   10575
         Left            =   -75000
         ScaleHeight     =   10515
         ScaleWidth      =   15555
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   15615
         Begin VB.CommandButton Command19 
            BackColor       =   &H00C0C0FF&
            Caption         =   "追加"
            Height          =   495
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   2520
            Width           =   855
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00C0C0FF&
            Caption         =   "锅号刷新"
            Height          =   495
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00C0C0FF&
            Caption         =   "工序删除"
            Height          =   495
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   1560
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "Formj1.frx":00D9
            Height          =   330
            Left            =   11400
            TabIndex        =   86
            Top             =   4680
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "工序编号"
            Text            =   "DataCombo2"
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00C0C0FF&
            Caption         =   "工序确定"
            Height          =   495
            Left            =   4560
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00C0C0FF&
            Caption         =   "全清"
            Height          =   495
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00C0C0FF&
            Caption         =   "全选"
            Height          =   495
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   1560
            Width           =   855
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            ItemData        =   "Formj1.frx":00EF
            Left            =   11040
            List            =   "Formj1.frx":00F1
            Style           =   1  'Checkbox
            TabIndex        =   72
            Top             =   1080
            Width           =   3735
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刷新"
            Height          =   495
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text13 
            Height          =   3375
            Left            =   11280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Text            =   "Formj1.frx":00F3
            Top             =   5160
            Width           =   3495
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00C0C0FF&
            Caption         =   "确认"
            Height          =   495
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   4560
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   4680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ListBox List3 
            Height          =   2580
            ItemData        =   "Formj1.frx":00FA
            Left            =   4800
            List            =   "Formj1.frx":00FC
            Style           =   1  'Checkbox
            TabIndex        =   67
            Top             =   5520
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7305
            ItemData        =   "Formj1.frx":00FE
            Left            =   6960
            List            =   "Formj1.frx":0100
            Style           =   1  'Checkbox
            TabIndex        =   66
            Top             =   960
            Width           =   3015
         End
         Begin MSAdodcLib.Adodc Adodc22 
            Height          =   330
            Left            =   2040
            Top             =   9000
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
         Begin MSAdodcLib.Adodc Adodc23 
            Height          =   330
            Left            =   2160
            Top             =   9000
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
         Begin MSAdodcLib.Adodc Adodc24 
            Height          =   330
            Left            =   2400
            Top             =   9000
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
         Begin MSAdodcLib.Adodc Adodc25 
            Height          =   330
            Left            =   4320
            Top             =   9000
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
         Begin MSAdodcLib.Adodc Adodc26 
            Height          =   330
            Left            =   6840
            Top             =   9000
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
         Begin MSAdodcLib.Adodc Adodc27 
            Height          =   330
            Left            =   7560
            Top             =   9000
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
            Bindings        =   "Formj1.frx":0102
            Height          =   7455
            Left            =   960
            TabIndex        =   76
            Top             =   1080
            Width           =   3015
            _cx             =   5318
            _cy             =   13150
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
            FormatString    =   $"Formj1.frx":0118
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFC0&
            Caption         =   "全选"
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
            Left            =   5880
            TabIndex        =   85
            Top             =   4680
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "工序信息"
            Height          =   495
            Index           =   3
            Left            =   11040
            TabIndex        =   84
            Top             =   600
            Width           =   3735
         End
         Begin VB.Label Label12 
            Caption         =   "工序内容"
            Height          =   3375
            Left            =   11040
            TabIndex        =   83
            Top             =   5160
            Width           =   255
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "工序编号"
            Height          =   495
            Index           =   4
            Left            =   11040
            TabIndex        =   82
            Top             =   4680
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "染色工序信息"
            Height          =   375
            Index           =   5
            Left            =   4560
            TabIndex        =   81
            Top             =   4200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "染色工序"
            Height          =   3015
            Index           =   13
            Left            =   4560
            TabIndex        =   80
            Top             =   5520
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFC0&
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
            Height          =   495
            Left            =   6360
            TabIndex        =   79
            Top             =   4680
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "所选工序"
            Height          =   375
            Index           =   7
            Left            =   6960
            TabIndex        =   78
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            Caption         =   "工序信息"
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   77
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   10575
         Left            =   -75000
         ScaleHeight     =   10515
         ScaleWidth      =   15555
         TabIndex        =   55
         Top             =   600
         Width           =   15615
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3210
            ItemData        =   "Formj1.frx":01ED
            Left            =   1320
            List            =   "Formj1.frx":01EF
            Style           =   1  'Checkbox
            TabIndex        =   228
            Top             =   2280
            Width           =   2055
         End
         Begin VB.CommandButton Command31 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刷新"
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Command30 
            BackColor       =   &H00C0C0FF&
            Caption         =   "确认"
            Height          =   375
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   226
            Top             =   1080
            Width           =   855
         End
         Begin MSAdodcLib.Adodc Adodc29 
            Height          =   330
            Left            =   3000
            Top             =   9720
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
            Caption         =   "Adodc29"
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
         Begin MSAdodcLib.Adodc Adodc28 
            Height          =   330
            Left            =   1800
            Top             =   9720
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
            Caption         =   "Adodc28"
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
            Left            =   12840
            TabIndex        =   24
            Text            =   "Text7"
            Top             =   4560
            Visible         =   0   'False
            Width           =   2420
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00C0C0FF&
            Caption         =   "返回"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00C0C0FF&
            Caption         =   "删除"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   3000
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "添加保存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10080
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2280
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   1575
            Left            =   11640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Text            =   "Formj1.frx":01F1
            Top             =   6000
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
            Bindings        =   "Formj1.frx":01F7
            Height          =   3375
            Left            =   5040
            TabIndex        =   56
            Top             =   2400
            Width           =   3375
            _cx             =   5953
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
            FormatString    =   $"Formj1.frx":020C
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
            Bindings        =   "Formj1.frx":02E1
            Height          =   330
            Index           =   16
            Left            =   12840
            TabIndex        =   23
            Top             =   4080
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "mc"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "Formj1.frx":02F6
            Height          =   360
            Left            =   12240
            TabIndex        =   209
            Top             =   2400
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "面料用途"
            Text            =   "DataCombo4"
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
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   330
            Left            =   12840
            TabIndex        =   22
            Top             =   2880
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   10
            Left            =   12480
            TabIndex        =   212
            Top             =   1920
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   13
            Left            =   12480
            TabIndex        =   213
            Top             =   1560
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   5880
            TabIndex        =   25
            Top             =   1200
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   309198849
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   5880
            TabIndex        =   26
            Top             =   1800
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   255
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   309198851
            CurrentDate     =   39177
         End
         Begin MSDataListLib.DataCombo DataCombo6 
            Bindings        =   "Formj1.frx":030C
            Height          =   330
            Left            =   12840
            TabIndex        =   220
            Top             =   5040
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "名称"
            Text            =   "DataCombo5"
         End
         Begin MSDataListLib.DataCombo DataCombo7 
            Bindings        =   "Formj1.frx":0322
            Height          =   330
            Left            =   12840
            TabIndex        =   221
            Top             =   5520
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "名称"
            Text            =   "DataCombo5"
         End
         Begin VB.Label Label1 
            Caption         =   "全选"
            Height          =   375
            Index           =   11
            Left            =   1320
            TabIndex        =   230
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "取消"
            Height          =   375
            Index           =   57
            Left            =   2520
            TabIndex        =   229
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "加工部门"
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
            Index           =   10
            Left            =   11640
            TabIndex        =   222
            Top             =   5520
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "生产方式"
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
            Index           =   9
            Left            =   11640
            TabIndex        =   219
            Top             =   5040
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "下单日期"
            Height          =   375
            Index           =   11
            Left            =   5040
            TabIndex        =   215
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "合同交期"
            Height          =   375
            Left            =   5040
            TabIndex        =   214
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "部门"
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
            Index           =   8
            Left            =   11640
            TabIndex        =   211
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "面料用途"
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
            Index           =   2
            Left            =   11640
            TabIndex        =   90
            Top             =   4560
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "投产类别"
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
            Index           =   1
            Left            =   11640
            TabIndex        =   60
            Top             =   4080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "备注"
            Height          =   495
            Index           =   8
            Left            =   11640
            TabIndex        =   58
            Top             =   2280
            Visible         =   0   'False
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   10455
         Left            =   0
         ScaleHeight     =   10395
         ScaleWidth      =   15555
         TabIndex        =   31
         Top             =   600
         Width           =   15615
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1440
            TabIndex        =   241
            Text            =   "Text11"
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CommandButton Command33 
            BackColor       =   &H00C0C0FF&
            Caption         =   "单号作废"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4080
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   239
            Top             =   1920
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DataCombo9 
            Bindings        =   "Formj1.frx":0338
            Height          =   330
            Left            =   12480
            TabIndex        =   233
            Top             =   3240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "单位"
            Text            =   "DataCombo9"
         End
         Begin VB.CommandButton Command32 
            BackColor       =   &H00C0C0FF&
            Caption         =   "工艺查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   231
            Top             =   4920
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   1
            Left            =   5280
            TabIndex        =   225
            Text            =   "Text1"
            Top             =   2880
            Width           =   615
         End
         Begin MSAdodcLib.Adodc Adodc30 
            Height          =   330
            Left            =   2400
            Top             =   9600
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
            Caption         =   "Adodc30"
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
         Begin VB.CommandButton Command29 
            BackColor       =   &H00C0C0FF&
            Caption         =   "合同打印"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13920
            Style           =   1  'Graphical
            TabIndex        =   216
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command26 
            BackColor       =   &H00C0C0FF&
            Caption         =   "计划"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   201
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "成品要求"
            Height          =   1215
            Left            =   5760
            TabIndex        =   98
            Top             =   1440
            Width           =   9735
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   330
               Index           =   4
               Left            =   360
               TabIndex        =   5
               Top             =   720
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   582
               _Version        =   393216
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   330
               Index           =   5
               Left            =   2160
               TabIndex        =   6
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "Formj1.frx":034E
               Height          =   330
               Index           =   22
               Left            =   4080
               TabIndex        =   7
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               ListField       =   "缩水率"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Bindings        =   "Formj1.frx":0363
               Height          =   330
               Index           =   23
               Left            =   6000
               TabIndex        =   8
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               ListField       =   "扭度"
               Text            =   "DataCombo1"
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   330
               Index           =   24
               Left            =   7920
               TabIndex        =   9
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               Text            =   "DataCombo1"
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "说明"
               Height          =   375
               Index           =   38
               Left            =   7920
               TabIndex        =   103
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "手感"
               Height          =   375
               Index           =   37
               Left            =   6000
               TabIndex        =   102
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "缩水"
               Height          =   375
               Index           =   36
               Left            =   4080
               TabIndex        =   101
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "克重"
               Height          =   375
               Index           =   6
               Left            =   2160
               TabIndex        =   100
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H0000C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "幅宽"
               Height          =   375
               Index           =   3
               Left            =   360
               TabIndex        =   99
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00C0C0FF&
            Caption         =   "报价查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13920
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo DataCombo3 
            Bindings        =   "Formj1.frx":0378
            Height          =   330
            Left            =   13920
            TabIndex        =   95
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "色别"
            Text            =   "DataCombo3"
         End
         Begin VB.TextBox Text1111 
            Height          =   270
            Left            =   2640
            TabIndex        =   62
            Text            =   "Text6"
            Top             =   6240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刷新"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   4920
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00C0C0FF&
            Caption         =   "生产打印"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   8760
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "删除"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "返回"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13320
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   4920
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "修改保存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4920
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "添加保存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3600
            MaskColor       =   &H00C0E0FF&
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4920
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0C0FF&
            Caption         =   "新单号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00C0C0FF&
            Caption         =   "查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   12720
            TabIndex        =   33
            Text            =   "Text2"
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   1320
            TabIndex        =   32
            Text            =   "Text4"
            Top             =   2880
            Width           =   735
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   6000
            Left            =   13680
            Top             =   720
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":038D
            Height          =   330
            Index           =   0
            Left            =   480
            TabIndex        =   1
            Top             =   3240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "简称"
            Text            =   "DataCombo1"
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Bindings        =   "Formj1.frx":03A2
            Height          =   3615
            Left            =   360
            TabIndex        =   37
            Top             =   5640
            Width           =   14895
            _cx             =   26273
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
            FormatString    =   $"Formj1.frx":03B8
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
         Begin MSAdodcLib.Adodc Adodc21 
            Height          =   330
            Left            =   7200
            Top             =   10920
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
            Caption         =   "Adodc21"
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
         Begin MSAdodcLib.Adodc Adodc20 
            Height          =   375
            Left            =   7320
            Top             =   10440
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
            Caption         =   "Adodc20"
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
         Begin MSAdodcLib.Adodc Adodc19 
            Height          =   330
            Left            =   7920
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
            Caption         =   "Adodc19"
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
         Begin MSAdodcLib.Adodc Adodc18 
            Height          =   375
            Left            =   7800
            Top             =   10440
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
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
            Caption         =   "Adodc18"
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
         Begin MSAdodcLib.Adodc Adodc17 
            Height          =   330
            Left            =   7920
            Top             =   10560
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
            Caption         =   "Adodc17"
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
         Begin MSAdodcLib.Adodc Adodc16 
            Height          =   375
            Left            =   7680
            Top             =   10680
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "Adodc16"
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
         Begin MSAdodcLib.Adodc Adodc15 
            Height          =   375
            Left            =   7320
            Top             =   10440
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "Adodc15"
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
         Begin MSAdodcLib.Adodc Adodc14 
            Height          =   330
            Left            =   7800
            Top             =   10560
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
            Caption         =   "Adodc14"
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
         Begin MSAdodcLib.Adodc Adodc13 
            Height          =   330
            Left            =   7800
            Top             =   10680
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
            Caption         =   "Adodc13"
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
         Begin MSAdodcLib.Adodc Adodc12 
            Height          =   330
            Left            =   8040
            Top             =   10560
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
            Caption         =   "Adodc12"
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
         Begin MSAdodcLib.Adodc Adodc11 
            Height          =   330
            Left            =   7560
            Top             =   10680
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
            Left            =   8040
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
            Height          =   375
            Left            =   7200
            Top             =   10440
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
            Height          =   375
            Left            =   7440
            Top             =   10440
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
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
            Height          =   375
            Left            =   8880
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
            Left            =   7080
            Top             =   10560
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
            Height          =   375
            Left            =   7200
            Top             =   10560
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
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
            Height          =   375
            Left            =   7440
            Top             =   10560
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
            Left            =   7800
            Top             =   10440
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
            Left            =   7560
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
            Left            =   8280
            Top             =   10560
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   1
            Left            =   1440
            TabIndex        =   42
            Top             =   1440
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   2
            Left            =   2160
            TabIndex        =   2
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":048D
            Height          =   330
            Index           =   3
            Left            =   3720
            TabIndex        =   3
            Top             =   3240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "布类"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   6
            Left            =   13320
            TabIndex        =   14
            Top             =   3240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   7
            Left            =   7680
            TabIndex        =   10
            Top             =   3240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   8
            Left            =   2280
            TabIndex        =   17
            Top             =   4560
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":04A2
            Height          =   330
            Index           =   9
            Left            =   9960
            TabIndex        =   20
            Top             =   3720
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   ""
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   11
            Left            =   14760
            TabIndex        =   18
            Top             =   4560
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":04B7
            Height          =   330
            Index           =   12
            Left            =   480
            TabIndex        =   16
            Top             =   4560
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "花型"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":04CC
            Height          =   330
            Index           =   15
            Left            =   2880
            TabIndex        =   43
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "mc"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   17
            Left            =   11520
            TabIndex        =   13
            Top             =   3240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":04E1
            Height          =   330
            Index           =   18
            Left            =   9120
            TabIndex        =   11
            Top             =   3240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "pm"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   19
            Left            =   3360
            TabIndex        =   19
            Top             =   4560
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":04F6
            Height          =   330
            Index           =   20
            Left            =   14520
            TabIndex        =   15
            Top             =   3240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "pm"
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   21
            Left            =   10440
            TabIndex        =   12
            Top             =   3240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   25
            Left            =   6000
            TabIndex        =   4
            Top             =   3240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   330
            Index           =   26
            Left            =   13800
            TabIndex        =   217
            Top             =   4560
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   582
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DataCombo8 
            Bindings        =   "Formj1.frx":050B
            Height          =   330
            Left            =   360
            TabIndex        =   224
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            _Version        =   393216
            Style           =   2
            ListField       =   "名称"
            Text            =   "DataCombo8"
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "Formj1.frx":0521
            Height          =   330
            Index           =   14
            Left            =   1440
            TabIndex        =   242
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "xm"
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "业务"
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
            Index           =   13
            Left            =   480
            TabIndex        =   243
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "缸号"
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
            Left            =   480
            TabIndex        =   240
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "单位"
            Height          =   375
            Index           =   58
            Left            =   12480
            TabIndex        =   232
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "布类"
            Height          =   255
            Index           =   56
            Left            =   360
            TabIndex        =   223
            Top             =   120
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "印花信息"
            Height          =   375
            Index           =   55
            Left            =   13800
            TabIndex        =   218
            Top             =   4200
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "批号"
            Height          =   375
            Index           =   54
            Left            =   6000
            TabIndex        =   210
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "色牢度"
            Height          =   375
            Index           =   32
            Left            =   10440
            TabIndex        =   97
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "颜色选择"
            Height          =   375
            Index           =   17
            Left            =   13920
            TabIndex        =   94
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "单价"
            Height          =   375
            Index           =   16
            Left            =   14520
            TabIndex        =   93
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "总备注"
            Height          =   375
            Index           =   15
            Left            =   3360
            TabIndex        =   92
            Top             =   4200
            Width           =   10335
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "色号"
            Height          =   375
            Index           =   14
            Left            =   9120
            TabIndex        =   89
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFC0&
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
            Left            =   4080
            TabIndex        =   64
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "匹数"
            Height          =   375
            Index           =   12
            Left            =   11520
            TabIndex        =   63
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "加工备注"
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   59
            Top             =   4200
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "序号"
            Height          =   375
            Index           =   9
            Left            =   14760
            TabIndex        =   54
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "染色要求"
            Height          =   375
            Index           =   7
            Left            =   8400
            TabIndex        =   53
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "颜色"
            Height          =   375
            Index           =   5
            Left            =   7680
            TabIndex        =   52
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "品名"
            Height          =   375
            Index           =   4
            Left            =   3720
            TabIndex        =   51
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "款号"
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   50
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "客户"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   49
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   漂 染 计 划 单"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   5400
            TabIndex        =   48
            Top             =   240
            Width           =   5655
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "来料单位"
            Height          =   375
            Left            =   480
            TabIndex        =   47
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "计划"
            Height          =   375
            Index           =   10
            Left            =   13320
            TabIndex        =   46
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "单号"
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
            Left            =   480
            TabIndex        =   45
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "代码"
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
            Left            =   11760
            TabIndex        =   44
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "Formj1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x As Integer
Dim BA As Database: Dim rr As Integer
Dim rs As Single: Dim RD1 As Recordset: Dim BA1 As Database: Public c, r As Integer: Public RQ As Date
Public mm As Date: Public ML As Date
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim ll As String
Dim cdbhf As Integer
''''''''''''''''''''''''''''''''''''''''''''''''
Dim sb As RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Private Sub Command12_Click()
If DataCombo1(11) = "" Then
If MsgBox("删除全部序号工序吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from dhgx  where 单号='" & DataCombo1(1) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Else
If MsgBox("删除序号" + DataCombo1(11) + "所有工序吗", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from dhgx  where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
DataCombo1(19) = ""
Adodc27.Refresh
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command11_Click()
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.Delete
Adodc3.Refresh
Text5.Text = ""
VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid2.ColWidth(2) = 4800
VSFlexGrid2.ColWidth(3) = 1200
VSFlexGrid2.ColWidth(4) = 1200
VSFlexGrid2.ColWidth(5) = 1200
End Sub

Private Sub Command13_Click()
On Error Resume Next
ll = ""
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
ll = ll + Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1) + "-"
End If
Next
Text13.Text = Mid(ll, 1, Len(ll) - 1)
End Sub

Private Sub Command14_Click()
Adodc22.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc22.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工序其它系数<>'0' and 工艺编号 not between '1001' and  '6000'  GROUP BY 工艺编号,工序名称 order by 工艺编号"
Adodc22.Refresh
If Adodc22.Recordset.EOF Then
List2.Clear
Exit Sub
End If
Adodc22.Recordset.MoveFirst
List2.Clear
Do While Not Adodc22.Recordset.EOF
List2.AddItem Adodc22.Recordset.Fields(0) + "-" + Trim(Adodc22.Recordset.Fields(1))
Adodc22.Recordset.MoveNext
Loop

End Sub

Private Sub Command15_Click()

Adodc26.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc26.RecordSource = "SELECT 工序 FROM DHGX WHERE 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "' order by 工序"
Adodc26.Refresh

If Not Adodc26.Recordset.EOF Then
Adodc26.Recordset.MoveFirst
ll = ""
Do While Not Adodc26.Recordset.EOF
ll = ll + Adodc26.Recordset.Fields(0) + "-"
Adodc26.Recordset.MoveNext
Loop
ll = Left(ll, Len(ll) - 1)
sql1 = "update kpd set gx='" & ll & "' where 单号='" & DataCombo1(1) & "' and ip='" & DataCombo1(11) & "'"
sql2 = "delete from ghgx where 锅号 in(select 锅号 from v_dh_gx where 单号= '" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "') and 工序 between '6001' and '9999' and 序号='" & DataCombo1(11) & "'"
sql3 = "insert into ghgx(锅号,序号,工序) select 锅号,序号,工序 from v_dh_gx where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
End If

End Sub

Private Sub Command16_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = True
Next
End Sub

Private Sub Command17_Click()
For i = 0 To List2.ListCount - 1
List2.Selected(i) = False
Next
End Sub

Private Sub Command18_Click()
If MsgBox("工序已选择，确认此类设置吗？", vbYesNo) = vbNo Then Exit Sub

If Text13.Text = "" Then
MsgBox ("请选择工序")
Exit Sub
End If
sclc = ""
For Q = 0 To List4.ListCount - 1
If List4.Selected(Q) = True Then
gxbh = Mid(List4.List(Q), 1, InStr(List4.List(Q), "-") - 1)
If Len(sclc) > 0 Then
sclc = sclc + "-" + Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1)
Else
sclc = sclc + Mid(List4.List(Q), InStr(List4.List(Q), "-") + 1)
End If
Adodc26.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc26.RecordSource = "select * from dhgx where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "' and 工序='" & gxbh & "'"
Adodc26.Refresh
If Adodc26.Recordset.EOF Then
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "dhgxlr('" & DataCombo1(1) & "','" & DataCombo1(11) & "','" & gxbh & "','1')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel
End If
End If
Next
Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.RecordSource = "select 序号,工序 from dhgx where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "' order by 序号,工序"
Adodc27.Refresh
MsgBox ("工序设置成功!")
If Len(sclc) > 0 Then
DataCombo1(19) = Mid(sclc, 1, Len(sclc) - 1)
End If
sql2 = "update sczy_x set 流程='" & DataCombo1(19) & "' where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End Sub

Private Sub Command19_Click()
If Text13.Text = "" Then
ll = Text13.Text
Else
ll = Text13.Text + "-"
End If
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
If InStr(ll, Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1)) = 0 Then
ll = ll + Mid(List2.List(i), 1, InStr(List2.List(i), "-") - 1) + "-"
End If
End If
Next
Text13.Text = Mid(ll, 1, Len(ll) - 1)
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub

If Len(DataCombo1(1)) = 0 Then
MsgBox ("订单编号有误，请确认!")
Exit Sub
End If

'If DataCombo8 = "" Then
'MsgBox ("请选择布类")
'Exit Sub
'End If

If DataCombo1(6).Text = "" Or DataCombo1(7).Text = "" Then
MsgBox ("输入不完整！")
Exit Sub
End If
DataCombo1(17) = Val(DataCombo1(17))
DataCombo1(20) = Val(DataCombo1(20))

For i = 0 To 26
Adodc18.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc18.Recordset.Fields(10) = DTPicker3.value
Adodc18.Recordset.Fields(13) = DTPicker4.value
Adodc18.Recordset.Fields(31) = DataCombo9
Adodc18.Recordset.Update
Adodc18.Refresh
DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1

Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh
DataCombo1(6).Text = ""
DataCombo1(7).Text = ""
DataCombo1(12).Text = ""
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command20_Click()
Formj11.Show
End Sub

Private Sub Command21_Click()
On Error Resume Next
If DataCombo1(1) = "" Then
MsgBox ("没有订单编号，禁止输入附注信息")
Exit Sub
End If

If MsgBox("确定保存附注信息吗？", vbYesNo) = vbNo Then Exit Sub

Adodc16.RecordSource = "select * from htfz_cpjf where 订单编号='" & DataCombo1(1) & "'"    ''''签约产品交付
Adodc16.Refresh
If Not Adodc16.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If

Adodc17.RecordSource = "select * from htfz_mlgg where 订单编号='" & DataCombo1(1) & "'"    ''''面料
Adodc17.Refresh
If Not Adodc17.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If

Adodc20.RecordSource = "select * from htfz_flgg where 订单编号='" & DataCombo1(1) & "'"    ''''辅料
Adodc20.Refresh
If Not Adodc20.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If

Adodc15.RecordSource = "select * from htfz_qtgg where 订单编号='" & DataCombo1(1) & "'"    ''''其它
Adodc15.Refresh
If Not Adodc15.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If


Adodc14.RecordSource = "select * from htfz_cpbmyq where 订单编号='" & DataCombo1(1) & "'"    ''''布面、责任、约定
Adodc14.Refresh
If Not Adodc14.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If

Adodc13.RecordSource = "select * from htfz_qybyj where 订单编号='" & DataCombo1(1) & "'"    ''''去向、验收、结算
Adodc13.Refresh
If Not Adodc13.Recordset.EOF Then
MsgBox ("附注已保存，点击附注显示，可以查询附注信息")
Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Adodc16.Recordset.AddNew     ''''签约产品交付
Adodc16.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 4
Adodc16.Recordset.Fields(i + 1) = Text14(i)
Next
Adodc16.Recordset.Update

Adodc17.Recordset.AddNew ''''面料
Adodc17.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc17.Recordset.Fields(i + 1) = Text24(i)
Next
Adodc17.Recordset.Update


Adodc20.Recordset.AddNew   ''''辅料
Adodc20.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc20.Recordset.Fields(i + 1) = Text25(i)
Next
Adodc20.Recordset.Update


Adodc15.Recordset.AddNew  ''''其它
Adodc15.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc15.Recordset.Fields(i + 1) = Text26(i)
Next
Adodc15.Recordset.Update


Adodc14.Recordset.AddNew ''''布面、责任、约定
Adodc14.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 5
Adodc14.Recordset.Fields(i + 1) = Text12(i)
Next
Adodc14.Recordset.Update

Adodc13.Recordset.AddNew ''''去向、验收、结算
Adodc13.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc13.Recordset.Fields(i + 1) = Text9(i)
Next
Adodc13.Recordset.Update

MsgBox ("保存成功！")
End Sub

Private Sub Command22_Click()
On Error Resume Next
Adodc16.RecordSource = "select * from htfz_cpjf where 订单编号='" & DataCombo1(1) & "'"    ''''签约产品交付
Adodc16.Refresh
If Adodc16.Recordset.EOF Then
For i = 0 To 4
Text14(i) = ""
Next
Else
For i = 0 To 3
Text14(i) = Adodc16.Recordset.Fields(i + 1)
Next
End If

Adodc17.RecordSource = "select * from htfz_mlgg where 订单编号='" & DataCombo1(1) & "'"    ''''面料
Adodc17.Refresh
If Adodc17.Recordset.EOF Then
For i = 0 To 8
Text24(i) = ""
Next
Else
For i = 0 To 8
Text24(i) = Adodc17.Recordset.Fields(i + 1)
Next
End If

Adodc20.RecordSource = "select * from htfz_flgg where 订单编号='" & DataCombo1(1) & "'"    ''''辅料
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
For i = 0 To 8
Text25(i) = ""
Next
Else
For i = 0 To 8
Text25(i) = Adodc20.Recordset.Fields(i + 1)
Next
End If


Adodc15.RecordSource = "select * from htfz_qtgg where 订单编号='" & DataCombo1(1) & "'"    ''''其它
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
For i = 0 To 8
Text26(i) = ""
Next
Else
For i = 0 To 8
Text26(i) = Adodc15.Recordset.Fields(i + 1)
Next
End If


Adodc14.RecordSource = "select * from htfz_cpbmyq where 订单编号='" & DataCombo1(1) & "'"    ''''布面、责任、约定
Adodc14.Refresh
If Adodc14.Recordset.EOF Then
For i = 0 To 5
Text12(i) = ""
Next
Else
For i = 0 To 5
Text12(i) = Adodc14.Recordset.Fields(i + 1)
Next
End If

Adodc13.RecordSource = "select * from htfz_qybyj where 订单编号='" & DataCombo1(1) & "'"    ''''去向、验收、结算
Adodc13.Refresh
If Adodc13.Recordset.EOF Then
For i = 0 To 8
Text9(i) = ""
Next
Else
For i = 0 To 8
Text9(i) = Adodc13.Recordset.Fields(i + 1)
Next
End If

End Sub

Private Sub Command23_Click()
On Error Resume Next
If DataCombo1(1) = "" Then
MsgBox ("没有订单编号，禁止输入附注信息")
Exit Sub
End If

If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub
     ''''签约产品交付
Adodc16.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 4
Adodc16.Recordset.Fields(i + 1) = Text14(i)
Next
Adodc16.Recordset.Update

 ''''面料
Adodc17.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc17.Recordset.Fields(i + 1) = Text24(i)
Next
Adodc17.Recordset.Update


   ''''辅料
Adodc20.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc20.Recordset.Fields(i + 1) = Text25(i)
Next
Adodc20.Recordset.Update


  ''''其它
Adodc15.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc15.Recordset.Fields(i + 1) = Text26(i)
Next
Adodc15.Recordset.Update


 ''''布面、责任、约定
Adodc14.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 5
Adodc14.Recordset.Fields(i + 1) = Text12(i)
Next
Adodc14.Recordset.Update

 ''''去向、验收、结算
Adodc13.Recordset.Fields(0) = DataCombo1(1)
For i = 0 To 8
Adodc13.Recordset.Fields(i + 1) = Text9(i)
Next
Adodc13.Recordset.Update

MsgBox ("修改成功！")

End Sub

Private Sub Command24_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc13.Recordset.Delete
Adodc14.Recordset.Delete
Adodc15.Recordset.Delete
Adodc16.Recordset.Delete
Adodc17.Recordset.Delete
Adodc20.Recordset.Delete
MsgBox ("删除成功1")
End Sub

Private Sub Command25_Click()
If Len(DataCombo1(1)) > 6 Then
Call htht(Adodc12, DataCombo1(1))
End If
End Sub

Private Sub Command26_Click()
Forma11.DataCombo1 = DataCombo1(0)
Forma11.DataCombo8 = DataCombo1(1)
Forma11.Show
End Sub

Private Sub Command27_Click()
If MsgBox("确定复制单号吗？", vbYesNo) = vbNo Then Exit Sub
If Len(DataCombo1(1)) = 10 Then
Adodc24.RecordSource = "select * from sczy_x where 单号='" & DataCombo1(1) & "'"
Adodc24.Refresh
If Not Adodc24.Recordset.EOF Then
MsgBox ("此订单号已经存在 禁止复制")
Exit Sub
End If

Adodc24.RecordSource = "select * from sczy_x where 单号='" & Text8(0) & "'"
Adodc24.Refresh
If Adodc24.Recordset.EOF Then
MsgBox ("被复制的单号不存在 不能复制")
Exit Sub
End If

If Text8(0) = DataCombo1(1) Then
MsgBox ("复制的单号不能是一个单号")
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''复制订单明细
sql1 = "INSERT INTO SCZY_X(客户,单号,款号,品名,成分,幅宽,克重,计划,色别,备注,技要,日期,序号,花型,交期,负责,排布,发货,匹数,色名,流程,单价,色牢度,缩水率,扭度,布纹) select 客户,'" & DataCombo1(1) & "',款号,品名,成分,幅宽,克重,计划,色别,备注,技要,日期,序号,花型,交期,负责,排布,发货,匹数,色名,流程,单价,色牢度,缩水率,扭度,布纹 from sczy_x  where 单号='" & Text8(0) & "'"
sql3 = "INSERT INTO dhgx(单号,序号,工序) select '" & DataCombo1(1) & "',序号,工序 from dhgx  where 单号='" & Text8(0) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
MsgBox ("单号复制成功！,请注意修改新单号的计划量！！")
Call Label5_Click
End If
End Sub

Private Sub Command28_Click()
If MsgBox("确定删除单号" + Text8(1) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from sczy_x  where 单号='" & Text8(1) & "'"
sql2 = "delete from sczy_z  where 单号='" & Text8(1) & "'"
sql3 = "delete from dhgx  where 单号='" & Text8(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
MsgBox ("删除成功！")
Call Label5_Click
End Sub

Private Sub Command29_Click()
Call XSHT(Adodc6, DataCombo1(1))
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command30_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql2 = "insert into dhjgxm(缸号,序号,项目,单价) VALUES('" & Text11 & "','" & DataCombo1(11) & "','" & List1.List(i) & "',0)"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
Next
Adodc3.RecordSource = "select 项目,缸号,序号 from dhjgxm where 缸号='" & Text11 & "' and 序号='" & DataCombo1(11) & "'"
Adodc3.Refresh
End Sub

Private Sub Command31_Click()
List1.Clear
Adodc2.RecordSource = "select 加工项目 from jgxm order by 序号"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
List1.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Loop
End If
List1.Selected(0) = True
Adodc3.RecordSource = "select 项目,缸号,序号 from dhjgxm where 缸号='" & Text11 & "' and 序号='" & DataCombo1(11) & "'"
Adodc3.Refresh
End Sub

Private Sub Command32_Click()
Formh224.DataCombo1(4) = DataCombo1(18)
Formh224.Check2(5).value = 1
Formh224.Show
End Sub

Private Sub Command33_Click()
Formh78.Text1 = DataCombo1(1)
Formh78.Show
End Sub

Private Sub Command34_Click()
On Error Resume Next
Formh79.Text1 = Text10
Formh79.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc18.Recordset.Delete
Adodc18.Refresh
DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1

Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh
DataCombo1(6).Text = ""
DataCombo1(7).Text = ""
DataCombo1(12).Text = ""
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
For i = 16 To 17
VSFlexGrid1.ColWidth(i) = 0
Next
End Sub

Private Sub Command5_Click()
'Adodc8.RecordSource = "select * from sczy_z where 计划锅号='" & Text7 & "'"
'Adodc8.Refresh
'If Adodc8.Recordset.EOF Then
'If DataCombo6 = "" Then
'MsgBox ("请选择加工类别！")
'Exit Sub
'End If
'If DataCombo7 = "" Then
'MsgBox ("请选择加工部门！")
'Exit Sub
'End If

If Adodc3.Recordset.EOF Then
Adodc3.Recordset.AddNew
Adodc3.Recordset.Fields(0) = DataCombo1(1).Text
Adodc3.Recordset.Fields(1) = Text5.Text
Adodc3.Recordset.Fields(2) = DataCombo1(16)
Adodc3.Recordset.Fields(3) = Text7
Adodc3.Recordset.Fields(4) = DataCombo5
Adodc3.Recordset.Fields(5) = DataCombo1(14)
Adodc3.Recordset.Fields(6) = DTPicker3.value
Adodc3.Recordset.Fields(7) = DTPicker4.value
Adodc3.Recordset.Fields(9) = DataCombo6
Adodc3.Recordset.Fields(10) = DataCombo7
Adodc3.Recordset.Update
Adodc3.Refresh
Else
Adodc3.Recordset.Fields(0) = DataCombo1(1).Text
Adodc3.Recordset.Fields(1) = Text5.Text
Adodc3.Recordset.Fields(2) = DataCombo1(16)
Adodc3.Recordset.Fields(3) = Text7
Adodc3.Recordset.Fields(4) = DataCombo5
Adodc3.Recordset.Fields(5) = DataCombo1(14)
Adodc3.Recordset.Fields(6) = DTPicker3.value
Adodc3.Recordset.Fields(7) = DTPicker4.value
Adodc3.Recordset.Fields(9) = DataCombo6
Adodc3.Recordset.Fields(10) = DataCombo7
Adodc3.Recordset.Update
Adodc3.Refresh
End If
VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid2.ColWidth(2) = 4600
VSFlexGrid2.ColWidth(3) = 1200
VSFlexGrid2.ColWidth(4) = 1200
VSFlexGrid2.ColWidth(5) = 1200
VSFlexGrid2.ColWidth(6) = 1200
VSFlexGrid2.ColWidth(7) = 1200
VSFlexGrid2.ColWidth(8) = 1200
End Sub

Private Sub Command6_Click()
Call SCTZD(Adodc6, DataCombo1(1))
End Sub

Private Sub Command7_Click()
On Error Resume Next
Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "select 单号  from v_sczy_x_dh where CONVERT(varchar(120),日期, 23)=CONVERT(varchar(120),GETDATE(), 23) and left(单号,1)='" & yhdm & "'  order by 单号 desc"
Adodc21.Refresh
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "001"
If Adodc21.Recordset.EOF Then
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "001"
Else
uu = Val(Right(Adodc21.Recordset.Fields(0), 3)) + 1
Select Case Len(uu)
       Case "1"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "00" + Trim(uu)
       Case "2"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "0" + Trim(uu)
       Case "3"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + Trim(uu)
End Select
End If

Adodc18.RecordSource = "select * from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc18.Refresh

Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc3.RecordSource = "select * from sczy_z where 单号='" & DataCombo1(1) & "'"
'Adodc3.Refresh

'Text1.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text7.Text = ""
'For i = 0 To 26
'DataCombo1(i).Text = ""
'Next
''''''''''''''''''''''''''''''''''''''附注信息
Adodc16.RecordSource = "select * from htfz_cpjf where 订单编号='" & DataCombo1(1) & "'"    ''''签约产品交付
Adodc16.Refresh
If Adodc16.Recordset.EOF Then
For i = 0 To 3
Text14(i) = ""
Next
Else
For i = 0 To 3
Text14(i) = Adodc16.Recordset.Fields(i + 1)
Next
End If

Adodc17.RecordSource = "select * from htfz_mlgg where 订单编号='" & DataCombo1(1) & "'"    ''''面料
Adodc17.Refresh
If Adodc17.Recordset.EOF Then
For i = 0 To 8
Text24(i) = ""
Next
Else
For i = 0 To 8
Text24(i) = Adodc17.Recordset.Fields(i + 1)
Next
End If

Adodc20.RecordSource = "select * from htfz_flgg where 订单编号='" & DataCombo1(1) & "'"    ''''辅料
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
For i = 0 To 8
Text24(i) = ""
Next
Else
For i = 0 To 8
Text24(i) = Adodc20.Recordset.Fields(i + 1)
Next
End If

Adodc20.RecordSource = "select * from htfz_flgg where 订单编号='" & DataCombo1(1) & "'"    ''''辅料
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
For i = 0 To 8
Text25(i) = ""
Next
Else
For i = 0 To 8
Text25(i) = Adodc20.Recordset.Fields(i + 1)
Next
End If

Adodc15.RecordSource = "select * from htfz_qtgg where 订单编号='" & DataCombo1(1) & "'"    ''''其它
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
For i = 0 To 8
Text26(i) = ""
Next
Else
For i = 0 To 8
Text26(i) = Adodc15.Recordset.Fields(i + 1)
Next
End If


Adodc14.RecordSource = "select * from htfz_cpbmyq where 订单编号='" & DataCombo1(1) & "'"    ''''布面、责任、约定
Adodc14.Refresh
If Adodc14.Recordset.EOF Then
For i = 0 To 5
Text12(i) = ""
Next
Else
For i = 0 To 5
Text12(i) = Adodc14.Recordset.Fields(i + 1)
Next
End If

Adodc13.RecordSource = "select * from htfz_qybyj where 订单编号='" & DataCombo1(1) & "'"    ''''去向、验收、结算
Adodc13.Refresh
If Adodc13.Recordset.EOF Then
For i = 0 To 8
Text9(i) = ""
Next
Else
For i = 0 To 8
Text9(i) = Adodc13.Recordset.Fields(i + 1)
Next
End If

'''''''''''''''''''''''''''''''

DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1
For i = 16 To 17
VSFlexGrid1.ColWidth(i) = 0
Next
End Sub

Private Sub Command1_Click()
On Error Resume Next


If Len(DataCombo1(1)) = 0 Then
MsgBox ("订单编号有误，请确认!")
Exit Sub
End If

If DataCombo9 = "" Then
MsgBox ("请输入单位!")
Exit Sub
End If

If DataCombo1(20).Text = "" Then DataCombo1(20).Text = 0
If DataCombo1(6).Text = "" Then DataCombo1(6) = 0
Adodc6.RecordSource = "select * from sczy_x where 客户='" & DataCombo1(0) & "' and 款号='" & DataCombo1(2) & "' and 品名='" & DataCombo1(3) & "' and 幅宽='" & DataCombo1(4) & "' and 克重='" & DataCombo1(5) & "' and 计划='" & DataCombo1(6) & "' and 色别='" & DataCombo1(7) & "' and 匹数='" & DataCombo1(17) & "' and 花型='" & DataCombo1(12) & "' and 备注='" & DataCombo1(8) & "'"
Adodc6.Refresh
If Not Adodc6.Recordset.EOF Then
If MsgBox("此信息已经输入，单号为" + Adodc6.Recordset.Fields(1) + "请确认是否继续输入？", vbYesNo) = vbNo Then Exit Sub
End If

Adodc6.RecordSource = "select 项目,单价 from dhjgxm where 缸号='" & Text11 & "' and 序号='" & DataCombo1(11) & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
If MsgBox("应该先设置加工项目，不设置也可以保存，不设置就保存吗？", vbYesNo) = vbNo Then Exit Sub
End If

DataCombo1(17) = Val(DataCombo1(17))
DataCombo1(20) = Val(DataCombo1(20))

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "YWSCZYX1('" & DataCombo1(0).Text & "','" & DataCombo1(1).Text & "','" & DataCombo1(2).Text & "','" & DataCombo1(3).Text & "','" & DataCombo1(4).Text & "','" & DataCombo1(5).Text & "','" & DataCombo1(6).Text & "','" & DataCombo1(7).Text & "','" & DataCombo1(8).Text & "','" & DataCombo1(9).Text & "','" & DTPicker3.value & "','" & DataCombo1(11).Text & "','" & DataCombo1(12).Text & "','" & DTPicker4.value & "','" & DataCombo1(14).Text & "','N','" & DataCombo1(16).Text & "','" & DataCombo1(17).Text & "','" & DataCombo1(18).Text & "','" & DataCombo1(19).Text & "','" & DataCombo1(20).Text & "','" & DataCombo1(21).Text & "','" & DataCombo1(22).Text & "','" & DataCombo1(23).Text & "','" & DataCombo1(24).Text & "','" & DataCombo1(25).Text & "','" & DataCombo1(26).Text & "','" & DataCombo8 & "','" & DataCombo9 & "','生产','" & Text11 & "')"       ' 表示调用哪个存储过程
    g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh

Adodc18.RecordSource = "SELECT * from SCZY_X where 单号= '" & DataCombo1(1).Text & "'  order by 序号"
Adodc18.Refresh
DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1
DataCombo1(6).Text = ""
DataCombo1(7).Text = ""
DataCombo1(12).Text = ""
End Sub

Private Sub Command8_Click()
On Error Resume Next
Timer1.Enabled = True

Adodc18.Refresh
Adodc7.Refresh
Adodc9.Refresh

Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh
DataCombo1(11).Text = 1
DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command9_Click()
Formj6.DTPicker1 = Date
Formj6.DTPicker2 = Date
Formj6.Check2(4).value = 1
Formj6.Show
End Sub



Private Sub DataCombo1_Change(Index As Integer)
Select Case Index
       Case 1
       
Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "select * from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc18.Refresh

DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1

Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh
Text11 = DataCombo1(1) + Trim(DataCombo1(11))

       Case 11
If DataCombo1(11) = "" Then
Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.RecordSource = "select 序号,工序 from dhgx where 单号='" & DataCombo1(1) & "' order by 序号,工序"
Adodc27.Refresh
Else
Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.RecordSource = "select 序号,工序 from dhgx where 单号='" & DataCombo1(1) & "' and 序号='" & DataCombo1(11) & "' order by 序号,工序"
Adodc27.Refresh
End If
Text11 = DataCombo1(1) + Trim(DataCombo1(11))
End Select

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 项目,缸号,序号 from dhjgxm where 缸号='" & Text11 & "' and 序号='" & DataCombo1(11) & "'"
Adodc3.Refresh

VSFlexGrid2.ColWidth(1) = 1200
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 1100
VSFlexGrid1.ColWidth(3) = 1100
VSFlexGrid1.ColWidth(4) = 2000
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub DataCombo2_Change()
If DataCombo2.Text = "" Then Exit Sub
Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "select 工序内容 from gybh  where 工序编号='" & DataCombo2.Text & "'"
Adodc25.Refresh
If Adodc25.Recordset.EOF Then
Text13.Text = ""
Else
Text13.Text = Adodc25.Recordset.Fields(0)
End If
End Sub

Private Sub DataCombo2_Click(Area As Integer)
If DataCombo2.Text = "" Then Exit Sub
Adodc25.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc25.RecordSource = "select 工序内容 from gybh  where 工序编号='" & DataCombo2.Text & "'"
Adodc25.Refresh
If Adodc25.Recordset.EOF Then
Text13.Text = ""
Else
Text13.Text = Adodc25.Recordset.Fields(0)
End If
End Sub

Private Sub DataCombo3_Change()
If Len(DataCombo3) > 0 Then
DataCombo1(7) = DataCombo3
End If
End Sub

Private Sub DataCombo3_Click(Area As Integer)
If Len(DataCombo3) > 0 Then
DataCombo1(7) = DataCombo3
End If
End Sub

Private Sub DataCombo4_Change()
Text7 = DataCombo4
End Sub

Private Sub DataCombo4_Click(Area As Integer)
Text7 = DataCombo4
End Sub

Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub



Private Sub DTPicker4_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub


Private Sub DTPicker5_Change()
Text14(4) = DTPicker5.value
End Sub

Private Sub DTPicker5_CloseUp()
Text14(4) = DTPicker5.value
End Sub

Private Sub Form_Load()
'
On Error Resume Next

cdbhf = cdbh


For i = 0 To 4
Text14(i) = ""
Text1(i) = ""
Next
Text11 = ""

For i = 0 To 5
Text12(i) = ""
Next

For i = 0 To 8
Text25(i) = ""
Text26(i) = ""
Text24(i) = ""
Next

For i = 0 To 8
Text9(i) = ""
Text8(i) = ""
Next
Text10 = ""
DataCombo9 = "公斤"
Text4.Text = ""
For i = 0 To 26
DataCombo1(i).Text = ""
Next
DTPicker3.value = Date
DTPicker1.value = Date
DTPicker2.value = Date
Text2.Text = ""
DTPicker4.value = Date
DTPicker5.value = Date
Text14(4) = Date
Text5.Text = ""
Text13.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo7.Text = ""
DataCombo8.Text = ""
Text7 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL WHERE IP LIKE '%'+'" & yhxx & "'+'%' group by 简称"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select distinct 扭度 from sczy_x"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select distinct 缩水率 from sczy_x"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select distinct 面料用途 from sczy_z"
Adodc10.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc19.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc19.RecordSource = "select distinct xm  from ywf"
Adodc19.Refresh
'If Not Adodc19.Recordset.EOF Then
'DataCombo1(14) = Adodc19.Recordset.Fields(0)
'End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select distinct 花型 from sczy_x"
Adodc5.Refresh

Adodc23.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc23.RecordSource = "select 工序编号,工序内容 from gybh "
Adodc23.Refresh

Adodc24.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc31.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc21.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc21.RecordSource = "select 单号  from v_sczy_x_dh where CONVERT(varchar(120),日期, 23)=CONVERT(varchar(120),GETDATE(), 23) and left(单号,1)='" & yhdm & "'  order by 单号 desc"
Adodc21.Refresh
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "001"
If Adodc21.Recordset.EOF Then
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "001"
Else
uu = Val(Right(Adodc21.Recordset.Fields(0), 3)) + 1
Select Case Len(uu)
       Case "1"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "00" + Trim(uu)
       Case "2"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + "0" + Trim(uu)
       Case "3"
DataCombo1(1).Text = yhdm + Format(DTPicker3.value, "YYMMDD") + Trim(uu)
End Select
End If

Adodc28.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc28.RecordSource = "select distinct 名称 from scfs"
Adodc28.Refresh

Adodc29.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc29.RecordSource = "select distinct 名称 from jgbm"
Adodc29.Refresh

Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select * from dhjgxm where 缸号='" & Text11 & "' and 序号='" & DataCombo1(11) & "'"
Adodc3.Refresh

Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "select * from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc18.Refresh
DataCombo1(12).Text = ""

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "select distinct 单位 from dddw"
Adodc12.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc20.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DataCombo6 = ""

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 1200
VSFlexGrid1.ColWidth(3) = 1200
VSFlexGrid1.ColWidth(4) = 2000

VSFlexGrid2.ColWidth(1) = 0
VSFlexGrid2.ColWidth(2) = 0
VSFlexGrid2.ColWidth(3) = 1200
VSFlexGrid2.ColWidth(4) = 0

If Len(yhdm) <> 1 Then
MsgBox ("这个账户不合适进入这个界面")
Command7.Enabled = False
Command8.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Label3.Enabled = False
End If

End Sub


Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(用户,菜单,编号) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where 用户='" & yhm & "' and 编号='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 8
Formj10.Show
       Case 11
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Select
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index

   Case 1
beizhu = 21
Forma112.Show
   
   Case 4
beizhu = 88
Forma17.Check1(0).value = 0
Forma17.Check1(2).value = 1
Forma17.DataCombo1 = DataCombo1(0)
Forma17.Show
   Case 5
   ysbl = 1
Forma38.Text1.Text = DataCombo1(7).Text
Forma38.Show
    Case 8
beizhu = 12
Forma112.Show

   Case 9
   DataCombo16.Enabled = True
   
   Case 15
   beizhu = 51
   Forma112.Text1(3) = DataCombo1(19)
   Forma112.Show

   Case 7
   beizhu = 52
   Forma112.Show
   
   Case 57
   For i = 0 To List1.ListCount - 1
   List1.Selected(i) = flase
   Next

   End Select
End Sub

Private Sub Label3_Click()
DataCombo1(1).Enabled = False
End Sub

Private Sub Label3_DblClick()
DataCombo1(1).Enabled = True
End Sub

Private Sub Label5_Click()
On Error Resume Next

Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "select * from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc18.Refresh

Adodc30.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc30.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 单号='" & DataCombo1(1).Text & "'  order by 序号"
Adodc30.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

DataCombo1(11).Text = Adodc18.Recordset.RecordCount + 1
Text11 = DataCombo1(1) + Trim(DataCombo1(11))

Adodc27.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc27.RecordSource = "select 序号,工序 from dhgx where 单号='" & DataCombo1(1) & "' order by 序号,工序"
Adodc27.Refresh

Adodc16.RecordSource = "select * from htfz_cpjf where 订单编号='" & DataCombo1(1) & "'"    ''''签约产品交付
Adodc16.Refresh
If Adodc16.Recordset.EOF Then
For i = 0 To 3
Text14(i) = ""
Next
Else
For i = 0 To 3
Text14(i) = Adodc16.Recordset.Fields(i + 1)
Next
End If

Adodc17.RecordSource = "select * from htfz_mlgg where 订单编号='" & DataCombo1(1) & "'"    ''''面料
Adodc17.Refresh
If Adodc17.Recordset.EOF Then
For i = 0 To 8
Text24(i) = ""
Next
Else
For i = 0 To 8
Text24(i) = Adodc17.Recordset.Fields(i + 1)
Next
End If

Adodc20.RecordSource = "select * from htfz_flgg where 订单编号='" & DataCombo1(1) & "'"    ''''辅料
Adodc20.Refresh
If Adodc20.Recordset.EOF Then
For i = 0 To 8
Text25(i) = ""
Next
Else
For i = 0 To 8
Text25(i) = Adodc20.Recordset.Fields(i + 1)
Next
End If


Adodc15.RecordSource = "select * from htfz_qtgg where 订单编号='" & DataCombo1(1) & "'"    ''''其它
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
For i = 0 To 8
Text26(i) = ""
Next
Else
For i = 0 To 8
Text26(i) = Adodc15.Recordset.Fields(i + 1)
Next
End If


Adodc14.RecordSource = "select * from htfz_cpbmyq where 订单编号='" & DataCombo1(1) & "'"    ''''布面、责任、约定
Adodc14.Refresh
If Adodc14.Recordset.EOF Then
For i = 0 To 5
Text12(i) = ""
Next
Else
For i = 0 To 5
Text12(i) = Adodc14.Recordset.Fields(i + 1)
Next
End If

Adodc13.RecordSource = "select * from htfz_qybyj where 订单编号='" & DataCombo1(1) & "'"    ''''去向、验收、结算
Adodc13.Refresh
If Adodc13.Recordset.EOF Then
For i = 0 To 8
Text9(i) = ""
Next
Else
For i = 0 To 8
Text9(i) = Adodc13.Recordset.Fields(i + 1)
Next
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 1100
VSFlexGrid1.ColWidth(3) = 1100
VSFlexGrid1.ColWidth(4) = 2000

End Sub


Private Sub Label6_Click()
Adodc31.RecordSource = "select 客户,单号,款号,品名,成分 as 批号,花型 as 货号,色别,色名 as 色号,幅宽,克重,匹数,单位,计划,备注 as 分备注,流程 as 总备注,缩水率 as 缩水,扭度 as 手感,布纹 as 说明,意见 as 印花,技要 as 染色要求,序号,日期,状态 from sczy_x where 状态='作废' order by 序号"
Adodc31.Refresh
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 1 Then
Call Command31_Click
End If
End Sub
Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case Index
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct 布类 from v_mp_kc where 客户名称='" & DataCombo1(0) & "' and 布类 like '%'+'" & Text1(0).Text & "' +'%' and 布类 like '%'+'" & Text1(1).Text & "' +'%'  order  by 布类"
Adodc4.Refresh
End Select
End Sub

Private Sub Text13_Change()
List4.Clear
i = 1
For L = 0 To Int(Len(Text13.Text) / 5)
''''''''''''''''''''''''''''''''''''''''''''''''''''
gxbh = Mid(Text13.Text, L * 4 + i, 4)
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 工艺编号,工序名称 from GYSHD where 工艺编号='" & gxbh & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
List4.AddItem Adodc8.Recordset.Fields(0) + "-" + Trim(Adodc8.Recordset.Fields(1))
End If
i = i + 1
Next
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End Sub

Private Sub Text4_Change()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 简称 from KHZL where 代码  like '%'+'" & Text4 & "'+'%' AND IP LIKE '%'+'" & yhxx & "'+'%' group by 简称"
Adodc1.Refresh
End Sub

Private Sub Timer1_Timer()
DTPicker3.value = Date
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next

rs = VSFlexGrid1.Row
If Adodc18.Recordset.EOF Then Exit Sub
Adodc18.Recordset.MoveFirst
Adodc18.Recordset.Move rs - 1
For i = 0 To 26
If i = 1 Then i = 2
DataCombo1(i).Text = Adodc18.Recordset.Fields(i)
Next
DTPicker3.value = Adodc18.Recordset.Fields(10)
DTPicker4.value = Adodc18.Recordset.Fields(13)
Adodc11.RecordSource = "select * from kpd where 单号='" & DataCombo1(1) & "' and ip='" & DataCombo1(11) & "'"
Adodc11.Refresh
If Not Adodc11.Recordset.EOF Then
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Else
Command1.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End If
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
rs = VSFlexGrid2.Row
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Move rs - 1
If MsgBox("删除" + "加工项目：" + Trim(Adodc3.Recordset.Fields(0)) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from dhjgxm  where 缸号='" & Adodc3.Recordset.Fields(2) & "' and 序号='" & Adodc3.Recordset.Fields(3) & "' and 项目='" & Adodc3.Recordset.Fields(0) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

Adodc3.Refresh
End Sub

Private Sub MSFlex()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
    If c <> 1 And c <> 2 And c <> 19 Then
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
    End If
End With
End Sub


Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlex
End If
End Sub

Private Sub text1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
Adodc30.Recordset.MoveFirst
Adodc30.Recordset.Move r - 1
Adodc30.Recordset.Fields(c - 1) = Text1111.Text
Adodc30.Recordset.Update
VSFlexGrid1.Text = Text1111.Text
Text1111.Visible = False
VSFlexGrid1.SetFocus
End If
End Sub



Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Not Adodc31.Recordset.EOF Then
Adodc31.Recordset.MoveFirst
rs = VSFlexGrid3.Row
Adodc31.Recordset.Move rs - 1
Text10 = Adodc31.Recordset.Fields(1)
End If
End Sub

Private Sub VSFlexGrid4_dblClick()
On Error Resume Next
If Adodc27.Recordset.EOF Then Exit Sub
rs = VSFlexGrid4.Row
Adodc27.Recordset.MoveFirst
Adodc27.Recordset.Move rs - 1
If MsgBox("删除" + "序号：" + Trim(Adodc27.Recordset.Fields(0)) + "工序：" + Adodc27.Recordset.Fields(1) + "吗？", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from dhgx  where 单号='" & DataCombo1(1) & "' and 序号='" & Adodc27.Recordset.Fields(0) & "' and 工序='" & Adodc27.Recordset.Fields(1) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc27.Refresh
End Sub

