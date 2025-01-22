VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "染整行业软件ERP系统"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   -690
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9915
   ScaleWidth      =   15720
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7200
      TabIndex        =   137
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formm1.frx":0000
      Height          =   390
      Left            =   1080
      TabIndex        =   136
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "菜单"
      Text            =   "DataCombo1"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   11895
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   24375
      _ExtentX        =   42995
      _ExtentY        =   20981
      _Version        =   393216
      Tabs            =   12
      Tab             =   9
      TabsPerRow      =   12
      TabHeight       =   1411
      BackColor       =   15904316
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "计划管理"
      TabPicture(0)   =   "Formm1.frx":0015
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture10"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "染缸计划"
      TabPicture(1)   =   "Formm1.frx":0031
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "毛坯管理"
      TabPicture(2)   =   "Formm1.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "化验管理"
      TabPicture(3)   =   "Formm1.frx":0069
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "对版管理"
      TabPicture(4)   =   "Formm1.frx":0085
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture4"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "生产管理"
      TabPicture(5)   =   "Formm1.frx":00A1
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture5(201)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "配料管理"
      TabPicture(6)   =   "Formm1.frx":00BD
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Picture7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "光坯管理"
      TabPicture(7)   =   "Formm1.frx":00D9
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Picture6"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "材料管理"
      TabPicture(8)   =   "Formm1.frx":00F5
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Picture8"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "财务管理"
      TabPicture(9)   =   "Formm1.frx":0111
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "Picture9"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "预警信息"
      TabPicture(10)  =   "Formm1.frx":012D
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Picture11"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "报表信息"
      TabPicture(11)  =   "Formm1.frx":0149
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Picture12"
      Tab(11).ControlCount=   1
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         Height          =   10935
         Index           =   201
         Left            =   -75600
         OLEDropMode     =   2  'Automatic
         ScaleHeight     =   10875
         ScaleWidth      =   24675
         TabIndex        =   51
         Top             =   840
         Width           =   24735
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "能耗统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   205
            Left            =   12360
            MouseIcon       =   "Formm1.frx":0165
            MousePointer    =   99  'Custom
            TabIndex        =   179
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "码单查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   100
            Left            =   12360
            MouseIcon       =   "Formm1.frx":167F7
            MousePointer    =   99  'Custom
            TabIndex        =   173
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   23
            Left            =   5160
            Picture         =   "Formm1.frx":2CE89
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   22
            Left            =   0
            Picture         =   "Formm1.frx":2D643
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "检测查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   131
            Left            =   1920
            MouseIcon       =   "Formm1.frx":2DDFD
            MousePointer    =   99  'Custom
            TabIndex        =   129
            Top             =   3000
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "物理检测"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   130
            Left            =   1920
            MouseIcon       =   "Formm1.frx":4448F
            MousePointer    =   99  'Custom
            TabIndex        =   128
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序分组"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   155
            Left            =   5760
            MouseIcon       =   "Formm1.frx":5AB21
            MousePointer    =   99  'Custom
            TabIndex        =   124
            Top             =   3000
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "委出查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   88
            Left            =   1920
            MouseIcon       =   "Formm1.frx":711B3
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   6000
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "委入查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   87
            Left            =   1920
            MouseIcon       =   "Formm1.frx":87845
            MousePointer    =   99  'Custom
            TabIndex        =   70
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "委外入库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   86
            Left            =   1920
            MouseIcon       =   "Formm1.frx":9DED7
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   4800
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "委外出库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   85
            Left            =   1920
            MouseIcon       =   "Formm1.frx":B4569
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   4200
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "部门工资"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   153
            Left            =   12360
            MouseIcon       =   "Formm1.frx":CABFB
            MousePointer    =   99  'Custom
            TabIndex        =   67
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "班组设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   140
            Left            =   5760
            MouseIcon       =   "Formm1.frx":E128D
            MousePointer    =   99  'Custom
            TabIndex        =   66
            Top             =   3720
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序扫描"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   118
            Left            =   12360
            MouseIcon       =   "Formm1.frx":F791F
            MousePointer    =   99  'Custom
            TabIndex        =   65
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "员工考勤"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   117
            Left            =   5760
            MouseIcon       =   "Formm1.frx":10DFB1
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   4440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   52
            Left            =   5760
            MouseIcon       =   "Formm1.frx":124643
            MousePointer    =   99  'Custom
            TabIndex        =   63
            Top             =   2280
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "报表明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   43
            Left            =   12360
            MouseIcon       =   "Formm1.frx":13ACD5
            MousePointer    =   99  'Custom
            TabIndex        =   62
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型码单"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   42
            Left            =   1920
            MouseIcon       =   "Formm1.frx":151367
            MousePointer    =   99  'Custom
            TabIndex        =   61
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "圆筒码单"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   41
            Left            =   1920
            MouseIcon       =   "Formm1.frx":1679F9
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   960
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   8
            Left            =   8160
            Picture         =   "Formm1.frx":17E08B
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "装卸统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   40
            Left            =   9120
            MouseIcon       =   "Formm1.frx":17E845
            MousePointer    =   99  'Custom
            TabIndex        =   59
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "质检统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   39
            Left            =   9120
            MouseIcon       =   "Formm1.frx":194ED7
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "开幅统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   38
            Left            =   9120
            MouseIcon       =   "Formm1.frx":1AB569
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "圆筒统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   37
            Left            =   9120
            MouseIcon       =   "Formm1.frx":1C1BFB
            MousePointer    =   99  'Custom
            TabIndex        =   56
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "烘干统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   36
            Left            =   9120
            MouseIcon       =   "Formm1.frx":1D828D
            MousePointer    =   99  'Custom
            TabIndex        =   55
            Top             =   3960
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "脱水统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   35
            Left            =   9120
            MouseIcon       =   "Formm1.frx":1EE91F
            MousePointer    =   99  'Custom
            TabIndex        =   54
            Top             =   4680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染色统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   34
            Left            =   9120
            MouseIcon       =   "Formm1.frx":204FB1
            MousePointer    =   99  'Custom
            TabIndex        =   53
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "排布统计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   33
            Left            =   9120
            MouseIcon       =   "Formm1.frx":21B643
            MousePointer    =   99  'Custom
            TabIndex        =   52
            Top             =   6120
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   0
            Left            =   11760
            Picture         =   "Formm1.frx":231CD5
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00FFFFFF&
         Height          =   11055
         Left            =   -75120
         ScaleHeight     =   10995
         ScaleWidth      =   24435
         TabIndex        =   120
         Top             =   840
         Width           =   24495
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "生产看板"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   204
            Left            =   3000
            MouseIcon       =   "Formm1.frx":23248F
            MousePointer    =   99  'Custom
            TabIndex        =   178
            Top             =   4440
            Width           =   2100
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "数据汇总表"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   203
            Left            =   3000
            MouseIcon       =   "Formm1.frx":248B21
            MousePointer    =   99  'Custom
            TabIndex        =   177
            Top             =   3360
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序履约明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   154
            Left            =   10920
            MouseIcon       =   "Formm1.frx":25F1B3
            MousePointer    =   99  'Custom
            TabIndex        =   131
            Top             =   4320
            Width           =   2610
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   21
            Left            =   9120
            Picture         =   "Formm1.frx":275845
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "光坯库存库龄"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   99
            Left            =   10920
            MouseIcon       =   "Formm1.frx":275FFF
            MousePointer    =   99  'Custom
            TabIndex        =   126
            Top             =   3120
            Width           =   2610
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00FFFFFF&
         Height          =   11055
         Left            =   0
         ScaleHeight     =   10995
         ScaleWidth      =   24195
         TabIndex        =   95
         Top             =   840
         Width           =   24255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "报价操作"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   91
            Left            =   7080
            MouseIcon       =   "Formm1.frx":28C691
            MousePointer    =   99  'Custom
            TabIndex        =   172
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "核算查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   165
            Left            =   9840
            TabIndex        =   145
            Top             =   5760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染助成本"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   164
            Left            =   9840
            TabIndex        =   144
            Top             =   1560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "成本明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   158
            Left            =   9840
            TabIndex        =   143
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "成本结余"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   133
            Left            =   9840
            TabIndex        =   142
            Top             =   4920
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   24
            Left            =   9360
            Picture         =   "Formm1.frx":2A2D23
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "成本分析"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   56
            Left            =   9840
            TabIndex        =   141
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "成本费用"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   13
            Left            =   9840
            TabIndex        =   140
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "车台看板"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   19
            Left            =   7080
            MouseIcon       =   "Formm1.frx":2A34DD
            MousePointer    =   99  'Custom
            TabIndex        =   139
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "染色单价"
            Height          =   375
            Index           =   92
            Left            =   13080
            TabIndex        =   119
            Top             =   9240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "期间设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   79
            Left            =   1080
            MouseIcon       =   "Formm1.frx":2B9B6F
            MousePointer    =   99  'Custom
            TabIndex        =   118
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "会计科目"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   80
            Left            =   1080
            MouseIcon       =   "Formm1.frx":2D0201
            MousePointer    =   99  'Custom
            TabIndex        =   117
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "摘要设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   81
            Left            =   1080
            MouseIcon       =   "Formm1.frx":2E6893
            MousePointer    =   99  'Custom
            TabIndex        =   116
            Top             =   2880
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "余额设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   82
            Left            =   1080
            MouseIcon       =   "Formm1.frx":2FCF25
            MousePointer    =   99  'Custom
            TabIndex        =   115
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "付票设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   83
            Left            =   1080
            MouseIcon       =   "Formm1.frx":3135B7
            MousePointer    =   99  'Custom
            TabIndex        =   114
            Top             =   4320
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "收票设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   84
            Left            =   1080
            MouseIcon       =   "Formm1.frx":329C49
            MousePointer    =   99  'Custom
            TabIndex        =   113
            Top             =   5040
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "付款未达"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   89
            Left            =   1080
            MouseIcon       =   "Formm1.frx":3402DB
            MousePointer    =   99  'Custom
            TabIndex        =   112
            Top             =   6480
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "员工信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   90
            Left            =   4320
            MouseIcon       =   "Formm1.frx":35696D
            MousePointer    =   99  'Custom
            TabIndex        =   111
            Top             =   2760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "供应资料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   93
            Left            =   4320
            MouseIcon       =   "Formm1.frx":36CFFF
            MousePointer    =   99  'Custom
            TabIndex        =   110
            Top             =   3480
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "客户资料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   94
            Left            =   4320
            MouseIcon       =   "Formm1.frx":383691
            MousePointer    =   99  'Custom
            TabIndex        =   109
            Top             =   4200
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "付票填制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   95
            Left            =   7080
            MouseIcon       =   "Formm1.frx":399D23
            MousePointer    =   99  'Custom
            TabIndex        =   108
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "收票填制"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   96
            Left            =   7080
            MouseIcon       =   "Formm1.frx":3B03B5
            MousePointer    =   99  'Custom
            TabIndex        =   107
            Top             =   4680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "特种记账"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   97
            Left            =   7080
            MouseIcon       =   "Formm1.frx":3C6A47
            MousePointer    =   99  'Custom
            TabIndex        =   106
            Top             =   3960
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "费用操作"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   98
            Left            =   7080
            MouseIcon       =   "Formm1.frx":3DD0D9
            MousePointer    =   99  'Custom
            TabIndex        =   105
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "对账明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   106
            Left            =   12480
            MouseIcon       =   "Formm1.frx":3F376B
            MousePointer    =   99  'Custom
            TabIndex        =   104
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "应付明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   107
            Left            =   12480
            MouseIcon       =   "Formm1.frx":409DFD
            MousePointer    =   99  'Custom
            TabIndex        =   103
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "收款未达"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   109
            Left            =   1080
            MouseIcon       =   "Formm1.frx":42048F
            MousePointer    =   99  'Custom
            TabIndex        =   102
            Top             =   5760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "费用查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   110
            Left            =   12480
            MouseIcon       =   "Formm1.frx":436B21
            MousePointer    =   99  'Custom
            TabIndex        =   101
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "应付查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   111
            Left            =   12480
            MouseIcon       =   "Formm1.frx":44D1B3
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   4320
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "应收查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   112
            Left            =   12480
            MouseIcon       =   "Formm1.frx":463845
            MousePointer    =   99  'Custom
            TabIndex        =   99
            Top             =   5040
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "特种查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   113
            Left            =   12480
            MouseIcon       =   "Formm1.frx":479ED7
            MousePointer    =   99  'Custom
            TabIndex        =   98
            Top             =   5760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "用户权限"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   114
            Left            =   4320
            MouseIcon       =   "Formm1.frx":490569
            MousePointer    =   99  'Custom
            TabIndex        =   97
            Top             =   4920
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料台账"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   115
            Left            =   12480
            MouseIcon       =   "Formm1.frx":4A6BFB
            MousePointer    =   99  'Custom
            TabIndex        =   96
            Top             =   2880
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   16
            Left            =   11880
            Picture         =   "Formm1.frx":4BD28D
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   18
            Left            =   6480
            Picture         =   "Formm1.frx":4BDA47
            Stretch         =   -1  'True
            Top             =   -360
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   19
            Left            =   3720
            Picture         =   "Formm1.frx":4BE201
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00FFFFFF&
         Height          =   11055
         Left            =   -75000
         ScaleHeight     =   10995
         ScaleWidth      =   24195
         TabIndex        =   81
         Top             =   840
         Width           =   24255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "退库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   208
            Left            =   11760
            MouseIcon       =   "Formm1.frx":4BE9BB
            MousePointer    =   99  'Custom
            TabIndex        =   182
            Top             =   4920
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料退库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   207
            Left            =   2400
            MouseIcon       =   "Formm1.frx":4D504D
            MousePointer    =   99  'Custom
            TabIndex        =   181
            Top             =   7440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "配料查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   11760
            MouseIcon       =   "Formm1.frx":4EB6DF
            MousePointer    =   99  'Custom
            TabIndex        =   94
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "车间设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   123
            Left            =   8280
            MouseIcon       =   "Formm1.frx":501D71
            MousePointer    =   99  'Custom
            TabIndex        =   93
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "总库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   78
            Left            =   11760
            MouseIcon       =   "Formm1.frx":518403
            MousePointer    =   99  'Custom
            TabIndex        =   92
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "出库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   77
            Left            =   11760
            MouseIcon       =   "Formm1.frx":52EA95
            MousePointer    =   99  'Custom
            TabIndex        =   91
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "入库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   76
            Left            =   11760
            MouseIcon       =   "Formm1.frx":545127
            MousePointer    =   99  'Custom
            TabIndex        =   90
            Top             =   1560
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8685
            Index           =   13
            Left            =   10320
            Picture         =   "Formm1.frx":55B7B9
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8565
            Index           =   11
            Left            =   7920
            Picture         =   "Formm1.frx":55BF73
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "库类设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   75
            Left            =   8280
            MouseIcon       =   "Formm1.frx":55C72D
            MousePointer    =   99  'Custom
            TabIndex        =   89
            Top             =   4440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "单位设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   74
            Left            =   8280
            MouseIcon       =   "Formm1.frx":572DBF
            MousePointer    =   99  'Custom
            TabIndex        =   88
            Top             =   3480
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "保管设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   72
            Left            =   8280
            MouseIcon       =   "Formm1.frx":589451
            MousePointer    =   99  'Custom
            TabIndex        =   87
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "盘存操作"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   71
            Left            =   2400
            MouseIcon       =   "Formm1.frx":59FAE3
            MousePointer    =   99  'Custom
            TabIndex        =   86
            Top             =   5040
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料盘存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   70
            Left            =   2400
            MouseIcon       =   "Formm1.frx":5B6175
            MousePointer    =   99  'Custom
            TabIndex        =   85
            Top             =   3840
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料出库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   69
            Left            =   2400
            MouseIcon       =   "Formm1.frx":5CC807
            MousePointer    =   99  'Custom
            TabIndex        =   84
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料入库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   68
            Left            =   2400
            MouseIcon       =   "Formm1.frx":5E2E99
            MousePointer    =   99  'Custom
            TabIndex        =   83
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "库存记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   67
            Left            =   2400
            MouseIcon       =   "Formm1.frx":5F952B
            MousePointer    =   99  'Custom
            TabIndex        =   82
            Top             =   6240
            Width           =   1740
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   18
            Left            =   3240
            Picture         =   "Formm1.frx":60FBBD
            Stretch         =   -1  'True
            Top             =   5640
            Width           =   345
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   17
            Left            =   3240
            Picture         =   "Formm1.frx":60FDD1
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   345
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   15
            Left            =   3240
            Picture         =   "Formm1.frx":60FFE5
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   345
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   12
            Left            =   3240
            Picture         =   "Formm1.frx":6101F9
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   345
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         Height          =   10815
         Left            =   -75000
         ScaleHeight     =   10755
         ScaleWidth      =   24195
         TabIndex        =   72
         Top             =   840
         Width           =   24255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "漏开查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   17
            Left            =   13680
            MouseIcon       =   "Formm1.frx":61040D
            MousePointer    =   99  'Custom
            TabIndex        =   171
            Top             =   6720
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "码单查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   161
            Left            =   18960
            MouseIcon       =   "Formm1.frx":626A9F
            MousePointer    =   99  'Custom
            TabIndex        =   132
            Top             =   2160
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "入库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   136
            Left            =   13680
            MouseIcon       =   "Formm1.frx":63D131
            MousePointer    =   99  'Custom
            TabIndex        =   80
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Image Image4 
            Height          =   405
            Index           =   0
            Left            =   4440
            Picture         =   "Formm1.frx":6537C3
            Stretch         =   -1  'True
            Top             =   4560
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "发货反审"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   135
            Left            =   5400
            MouseIcon       =   "Formm1.frx":653A0F
            MousePointer    =   99  'Custom
            TabIndex        =   79
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Image Image5 
            Height          =   540
            Index           =   3
            Left            =   3120
            Picture         =   "Formm1.frx":66A0A1
            Stretch         =   -1  'True
            Top             =   3600
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "出库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   50
            Left            =   13680
            MouseIcon       =   "Formm1.frx":66A2B5
            MousePointer    =   99  'Custom
            TabIndex        =   78
            Top             =   3480
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "库存查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   49
            Left            =   13680
            MouseIcon       =   "Formm1.frx":680947
            MousePointer    =   99  'Custom
            TabIndex        =   77
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "发货报表"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   48
            Left            =   13680
            MouseIcon       =   "Formm1.frx":696FD9
            MousePointer    =   99  'Custom
            TabIndex        =   76
            Top             =   5640
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "光坯库存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   47
            Left            =   2520
            MouseIcon       =   "Formm1.frx":6AD66B
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   6360
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "光坯发货"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   46
            Left            =   2520
            MouseIcon       =   "Formm1.frx":6C3CFD
            MousePointer    =   99  'Custom
            TabIndex        =   74
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "光坯入库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   45
            Left            =   2520
            MouseIcon       =   "Formm1.frx":6DA38F
            MousePointer    =   99  'Custom
            TabIndex        =   73
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   9525
            Index           =   2
            Left            =   10440
            Picture         =   "Formm1.frx":6F0A21
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Image Image5 
            Height          =   540
            Index           =   20
            Left            =   3120
            Picture         =   "Formm1.frx":6F11DB
            Stretch         =   -1  'True
            Top             =   5640
            Width           =   465
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   10815
         Left            =   -75000
         ScaleHeight     =   10755
         ScaleWidth      =   24195
         TabIndex        =   44
         Top             =   840
         Width           =   24255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "并缸查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   206
            Left            =   12360
            MouseIcon       =   "Formm1.frx":6F13EF
            MousePointer    =   99  'Custom
            TabIndex        =   180
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "料单成本"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   201
            Left            =   12360
            MouseIcon       =   "Formm1.frx":707A81
            MousePointer    =   99  'Custom
            TabIndex        =   175
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染色时长"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   162
            Left            =   7680
            MouseIcon       =   "Formm1.frx":71E113
            MousePointer    =   99  'Custom
            TabIndex        =   133
            Top             =   5520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "缸号设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   143
            Left            =   7680
            MouseIcon       =   "Formm1.frx":7347A5
            MousePointer    =   99  'Custom
            TabIndex        =   130
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染色工序"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   7680
            MouseIcon       =   "Formm1.frx":74AE37
            MousePointer    =   99  'Custom
            TabIndex        =   125
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染色确认"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   63
            Left            =   12360
            MouseIcon       =   "Formm1.frx":7614C9
            MousePointer    =   99  'Custom
            TabIndex        =   50
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   7
            Left            =   10440
            Picture         =   "Formm1.frx":777B5B
            Stretch         =   -1  'True
            Top             =   -360
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   6
            Left            =   6960
            Picture         =   "Formm1.frx":778315
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "投染信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   32
            Left            =   12360
            MouseIcon       =   "Formm1.frx":778ACF
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "配料信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   31
            Left            =   12360
            MouseIcon       =   "Formm1.frx":78F161
            MousePointer    =   99  'Custom
            TabIndex        =   48
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "操作员工"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   30
            Left            =   7680
            MouseIcon       =   "Formm1.frx":7A57F3
            MousePointer    =   99  'Custom
            TabIndex        =   47
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工艺设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   29
            Left            =   7680
            MouseIcon       =   "Formm1.frx":7BBE85
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "生成配料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   28
            Left            =   2400
            MouseIcon       =   "Formm1.frx":7D2517
            MousePointer    =   99  'Custom
            TabIndex        =   45
            Top             =   3360
            Width           =   1740
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   10815
         Left            =   -75000
         ScaleHeight     =   10755
         ScaleWidth      =   24075
         TabIndex        =   32
         Top             =   840
         Width           =   24135
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "模板设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   64
            Left            =   8760
            MouseIcon       =   "Formm1.frx":7E8BA9
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "复样打样查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   119
            Left            =   11880
            MouseIcon       =   "Formm1.frx":7FF23B
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   6000
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "客样打样查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   152
            Left            =   11880
            MouseIcon       =   "Formm1.frx":8158CD
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   4800
            Width           =   2610
         End
         Begin VB.Image Image5 
            Height          =   660
            Index           =   13
            Left            =   2280
            Picture         =   "Formm1.frx":82BF5F
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "客样管理"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   132
            Left            =   1680
            MouseIcon       =   "Formm1.frx":82C173
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   960
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "大货配方查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   27
            Left            =   11880
            MouseIcon       =   "Formm1.frx":842805
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   3360
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "确认配方查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   26
            Left            =   11880
            MouseIcon       =   "Formm1.frx":858E97
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Top             =   1920
            Width           =   2610
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   5
            Left            =   8400
            Picture         =   "Formm1.frx":86F529
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   4
            Left            =   10920
            Picture         =   "Formm1.frx":86FCE3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "化验设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   25
            Left            =   8760
            MouseIcon       =   "Formm1.frx":87049D
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "色号报价"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   24
            Left            =   5280
            MouseIcon       =   "Formm1.frx":886B2F
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   3120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "大货配方"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   23
            Left            =   1680
            MouseIcon       =   "Formm1.frx":89D1C1
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   5760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "确认配方"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   22
            Left            =   1680
            MouseIcon       =   "Formm1.frx":8B3853
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   3120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "配方设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   21
            Left            =   8760
            MouseIcon       =   "Formm1.frx":8C9EE5
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Image Image4 
            Height          =   405
            Index           =   3
            Left            =   4080
            Picture         =   "Formm1.frx":8E0577
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   690
         End
         Begin VB.Image Image5 
            Height          =   660
            Index           =   1
            Left            =   2280
            Picture         =   "Formm1.frx":8E07C3
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   345
         End
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   10815
         Left            =   -75000
         Picture         =   "Formm1.frx":8E09D7
         ScaleHeight     =   10755
         ScaleWidth      =   24075
         TabIndex        =   24
         Top             =   840
         Width           =   24135
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "毛坯明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   202
            Left            =   11640
            MouseIcon       =   "Formm1.frx":8F7F79
            MousePointer    =   99  'Custom
            TabIndex        =   176
            Top             =   6120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "退库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   62
            Left            =   11640
            MouseIcon       =   "Formm1.frx":90E60B
            MousePointer    =   99  'Custom
            TabIndex        =   169
            Top             =   5280
            Width           =   1740
         End
         Begin VB.Image Image6 
            Height          =   615
            Index           =   4
            Left            =   3360
            Picture         =   "Formm1.frx":924C9D
            Stretch         =   -1  'True
            Top             =   2040
            Width           =   375
         End
         Begin VB.Image Image6 
            Height          =   615
            Index           =   3
            Left            =   3360
            Picture         =   "Formm1.frx":924EB1
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "库存记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   122
            Left            =   2640
            MouseIcon       =   "Formm1.frx":9250C5
            MousePointer    =   99  'Custom
            TabIndex        =   168
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "毛坯库存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   121
            Left            =   2640
            MouseIcon       =   "Formm1.frx":93B757
            MousePointer    =   99  'Custom
            TabIndex        =   167
            Top             =   2760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "码单查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   57
            Left            =   11640
            MouseIcon       =   "Formm1.frx":951DE9
            MousePointer    =   99  'Custom
            TabIndex        =   134
            Top             =   4320
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "毛坯退库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   3
            Left            =   2640
            MouseIcon       =   "Formm1.frx":96847B
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Image Image6 
            Height          =   615
            Index           =   0
            Left            =   3360
            Picture         =   "Formm1.frx":97EB0D
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   375
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   10
            Left            =   6960
            Picture         =   "Formm1.frx":97ED21
            Stretch         =   -1  'True
            Top             =   -360
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   9
            Left            =   10080
            Picture         =   "Formm1.frx":97F4DB
            Stretch         =   -1  'True
            Top             =   -480
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "库存查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   16
            Left            =   11640
            MouseIcon       =   "Formm1.frx":97FC95
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "备活报表"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   11640
            MouseIcon       =   "Formm1.frx":996327
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   2400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "入库查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   11640
            MouseIcon       =   "Formm1.frx":9AC9B9
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "布类设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   7
            Left            =   7680
            MouseIcon       =   "Formm1.frx":9C304B
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "保管设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   6
            Left            =   7680
            MouseIcon       =   "Formm1.frx":9D96DD
            MousePointer    =   99  'Custom
            TabIndex        =   26
            Top             =   2280
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "毛坯入库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   2640
            MouseIcon       =   "Formm1.frx":9EFD6F
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   5640
            Width           =   1740
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   10935
         Left            =   -75000
         ScaleHeight     =   10875
         ScaleWidth      =   24075
         TabIndex        =   21
         Top             =   840
         Width           =   24135
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "车台看板"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   200
            Left            =   3120
            MouseIcon       =   "Formm1.frx":A06401
            MousePointer    =   99  'Custom
            TabIndex        =   174
            Top             =   5280
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "排缸查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   44
            Left            =   11760
            MouseIcon       =   "Formm1.frx":A1CA93
            MousePointer    =   99  'Custom
            TabIndex        =   123
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   1
            Left            =   10080
            Picture         =   "Formm1.frx":A33125
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "车台设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   18
            Left            =   8040
            MouseIcon       =   "Formm1.frx":A338DF
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染缸计划"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   20
            Left            =   3120
            MouseIcon       =   "Formm1.frx":A49F71
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   3
            Left            =   7680
            Picture         =   "Formm1.frx":A60603
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00FFFFFF&
         Height          =   10935
         Left            =   -75000
         ScaleHeight     =   10875
         ScaleWidth      =   22635
         TabIndex        =   18
         Top             =   840
         Width           =   22695
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "日志查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   15
            Left            =   12240
            MouseIcon       =   "Formm1.frx":A60DBD
            MousePointer    =   99  'Custom
            TabIndex        =   170
            Top             =   4920
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "调度设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   73
            Left            =   8640
            MouseIcon       =   "Formm1.frx":A7744F
            MousePointer    =   99  'Custom
            TabIndex        =   166
            Top             =   2760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染耗查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   14
            Left            =   12240
            MouseIcon       =   "Formm1.frx":A8DAE1
            MousePointer    =   99  'Custom
            TabIndex        =   165
            Top             =   3840
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "布类设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   126
            Left            =   8640
            MouseIcon       =   "Formm1.frx":AA4173
            MousePointer    =   99  'Custom
            TabIndex        =   127
            Top             =   6000
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "计划查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   11
            Left            =   12240
            MouseIcon       =   "Formm1.frx":ABA805
            MousePointer    =   99  'Custom
            TabIndex        =   122
            Top             =   2760
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "生产计划"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   9
            Left            =   2400
            MouseIcon       =   "Formm1.frx":AD0E97
            MousePointer    =   99  'Custom
            TabIndex        =   121
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "备注设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   8640
            MouseIcon       =   "Formm1.frx":AE7529
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   3840
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "业务设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   103
            Left            =   8640
            MouseIcon       =   "Formm1.frx":AFDBBB
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   4920
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   15
            Left            =   7680
            Picture         =   "Formm1.frx":B1424D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   14
            Left            =   11040
            Picture         =   "Formm1.frx":B14A07
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H00FFFFFF&
         Height          =   11055
         Left            =   -75000
         ScaleHeight     =   10995
         ScaleWidth      =   24195
         TabIndex        =   9
         Top             =   840
         Width           =   24255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "欠费预警信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   151
            Left            =   10560
            MouseIcon       =   "Formm1.frx":B151C1
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   5280
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料预警信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   150
            Left            =   10560
            MouseIcon       =   "Formm1.frx":B2B853
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   4320
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料预警信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   149
            Left            =   10560
            MouseIcon       =   "Formm1.frx":B41EE5
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   3360
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序预警信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   147
            Left            =   10560
            MouseIcon       =   "Formm1.frx":B58577
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   2400
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "工序预警设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   142
            Left            =   3000
            MouseIcon       =   "Formm1.frx":B6EC09
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   2400
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料预警设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   139
            Left            =   3000
            MouseIcon       =   "Formm1.frx":B8529B
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   3360
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料预警设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   137
            Left            =   3000
            MouseIcon       =   "Formm1.frx":B9B92D
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   4320
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "欠费预警设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   138
            Left            =   3000
            MouseIcon       =   "Formm1.frx":BB1FBF
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   5280
            Width           =   2610
         End
         Begin VB.Image Image8 
            Height          =   8205
            Index           =   17
            Left            =   8280
            Picture         =   "Formm1.frx":BC8651
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         Height          =   10935
         Left            =   -75000
         ScaleHeight     =   10875
         ScaleWidth      =   24060
         TabIndex        =   1
         Top             =   840
         Width           =   24120
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型配方"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   101
            Left            =   9360
            MouseIcon       =   "Formm1.frx":BC8E0B
            MousePointer    =   99  'Custom
            TabIndex        =   164
            Top             =   1920
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型工艺"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   51
            Left            =   9360
            MouseIcon       =   "Formm1.frx":BDF49D
            MousePointer    =   99  'Custom
            TabIndex        =   163
            Top             =   1320
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型助剂"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   166
            Left            =   9360
            MouseIcon       =   "Formm1.frx":BF5B2F
            MousePointer    =   99  'Custom
            TabIndex        =   162
            Top             =   720
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型机台"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   9360
            MouseIcon       =   "Formm1.frx":C0C1C1
            MousePointer    =   99  'Custom
            TabIndex        =   161
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Image Image4 
            Height          =   405
            Index           =   4
            Left            =   4200
            Picture         =   "Formm1.frx":C22853
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   690
         End
         Begin VB.Image Image4 
            Height          =   405
            Index           =   2
            Left            =   4200
            Picture         =   "Formm1.frx":C22A9F
            Stretch         =   -1  'True
            Top             =   5400
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料储备"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   177
            Left            =   1800
            MouseIcon       =   "Formm1.frx":C22CEB
            MousePointer    =   99  'Custom
            TabIndex        =   160
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "粉体储备"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   176
            Left            =   1800
            MouseIcon       =   "Formm1.frx":C3937D
            MousePointer    =   99  'Custom
            TabIndex        =   159
            Top             =   4320
            Width           =   1740
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   2
            Left            =   6360
            Picture         =   "Formm1.frx":C4FA0F
            Stretch         =   -1  'True
            Top             =   6120
            Width           =   345
         End
         Begin VB.Image Image6 
            Height          =   495
            Index           =   2
            Left            =   6360
            Picture         =   "Formm1.frx":C4FC23
            Stretch         =   -1  'True
            Top             =   4920
            Width           =   375
         End
         Begin VB.Image Image6 
            Height          =   495
            Index           =   1
            Left            =   6360
            Picture         =   "Formm1.frx":C4FE37
            Stretch         =   -1  'True
            Top             =   3720
            Width           =   375
         End
         Begin VB.Image Image5 
            Height          =   420
            Index           =   0
            Left            =   6360
            Picture         =   "Formm1.frx":C5004B
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   345
         End
         Begin VB.Image Image5 
            Height          =   540
            Index           =   4
            Left            =   6360
            Picture         =   "Formm1.frx":C5025F
            Stretch         =   -1  'True
            Top             =   2400
            Width           =   345
         End
         Begin VB.Image Image4 
            Height          =   405
            Index           =   1
            Left            =   4200
            Picture         =   "Formm1.frx":C50473
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "定型申请"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   1800
            MouseIcon       =   "Formm1.frx":C506BF
            MousePointer    =   99  'Custom
            TabIndex        =   158
            Top             =   1680
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "配料审核"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   156
            Left            =   1800
            MouseIcon       =   "Formm1.frx":C66D51
            MousePointer    =   99  'Custom
            TabIndex        =   157
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "助剂申请"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   170
            Left            =   5760
            MouseIcon       =   "Formm1.frx":C7D3E3
            MousePointer    =   99  'Custom
            TabIndex        =   156
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   171
            Left            =   5760
            MouseIcon       =   "Formm1.frx":C93A75
            MousePointer    =   99  'Custom
            TabIndex        =   155
            Top             =   6720
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送监控"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   173
            Left            =   5760
            MouseIcon       =   "Formm1.frx":CAA107
            MousePointer    =   99  'Custom
            TabIndex        =   154
            Top             =   3120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "液体染缸"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   129
            Left            =   9360
            MouseIcon       =   "Formm1.frx":CC0799
            MousePointer    =   99  'Custom
            TabIndex        =   153
            Top             =   3600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送助剂"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   167
            Left            =   9360
            MouseIcon       =   "Formm1.frx":CD6E2B
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   6600
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送粉体"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   168
            Left            =   9360
            MouseIcon       =   "Formm1.frx":CED4BD
            MousePointer    =   99  'Custom
            TabIndex        =   151
            Top             =   6000
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "粉体染缸"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   174
            Left            =   9360
            MouseIcon       =   "Formm1.frx":D03B4F
            MousePointer    =   99  'Custom
            TabIndex        =   150
            Top             =   4200
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料染缸"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   175
            Left            =   9360
            MouseIcon       =   "Formm1.frx":D1A1E1
            MousePointer    =   99  'Custom
            TabIndex        =   149
            Top             =   4800
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8925
            Index           =   25
            Left            =   4560
            Picture         =   "Formm1.frx":D30873
            Stretch         =   -1  'True
            Top             =   -240
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料输送"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   59
            Left            =   5760
            MouseIcon       =   "Formm1.frx":D3102D
            MousePointer    =   99  'Custom
            TabIndex        =   148
            Top             =   5520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "粉体输送"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   58
            Left            =   5760
            MouseIcon       =   "Formm1.frx":D476BF
            MousePointer    =   99  'Custom
            TabIndex        =   147
            Top             =   4320
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "助剂输送"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   53
            Left            =   5760
            MouseIcon       =   "Formm1.frx":D5DD51
            MousePointer    =   99  'Custom
            TabIndex        =   146
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8925
            Index           =   26
            Left            =   8520
            Picture         =   "Formm1.frx":D743E3
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "高浓设置"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   120
            Left            =   12360
            MouseIcon       =   "Formm1.frx":D74B9D
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染助料库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   66
            Left            =   12360
            MouseIcon       =   "Formm1.frx":D8B22F
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   2280
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "报价料库"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   65
            Left            =   12360
            MouseIcon       =   "Formm1.frx":DA18C1
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3960
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "染料称量"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   61
            Left            =   1800
            MouseIcon       =   "Formm1.frx":DB7F53
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   6120
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "助剂称量"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   60
            Left            =   1800
            MouseIcon       =   "Formm1.frx":DCE5E5
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   6840
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "称量染料"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   440
            Index           =   55
            Left            =   1800
            MouseIcon       =   "Formm1.frx":DE4C77
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "称量液体"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   440
            Index           =   54
            Left            =   1800
            MouseIcon       =   "Formm1.frx":DFB309
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Image Image8 
            Height          =   8925
            Index           =   12
            Left            =   11760
            Picture         =   "Formm1.frx":E1199B
            Stretch         =   -1  'True
            Top             =   -240
            Width           =   15
         End
      End
      Begin VB.Image Image2 
         Height          =   8385
         Index           =   0
         Left            =   -75000
         Picture         =   "Formm1.frx":E12155
         Stretch         =   -1  'True
         Top             =   840
         Width           =   15405
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9840
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   15720
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Label Label3 
      Caption         =   "菜单编号"
      Height          =   375
      Left            =   6240
      TabIndex        =   138
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "运行菜单"
      Height          =   375
      Left            =   120
      TabIndex        =   135
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   0
      Picture         =   "Formm1.frx":E1A093
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "Formm1.frx":E1AD5D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15210
   End
   Begin VB.Menu GSWZ 
      Caption         =   "公司网站"
   End
   Begin VB.Menu BZWD 
      Caption         =   "帮助文档"
   End
   Begin VB.Menu TCXT 
      Caption         =   "退出系统"
   End
End
Attribute VB_Name = "Formm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cdsx As Integer
Private Sub BZWD_Click()
SendKeys "{F1}"
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
If DataCombo1 = "" Then Exit Sub
Adodc2.RecordSource = "select distinct 编号 from yhcd where 用户='" & yhm & "' and 菜单='" & DataCombo1 & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Call zqcd(Adodc2.Recordset.Fields(0))
End If
DataCombo1 = ""
End Sub

Private Sub Form_Load()
DataCombo1 = ""
cdsx = 1
Text1 = ""
xtxxjm = "更新"
App.HelpFile = App.Path & "\help.chm"
Formm1.Caption = ljb
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 菜单,编号 from yhcd where  用户='" & yhm & "'"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub GSWZ_Click()
Call ShellExecute(Me.hwnd, "open", "http://www.wffjrj.com", vbNullString, vbNullString, &H0)
End Sub

Private Sub Label1_Click(Index As Integer)

Text1 = Index
cdbh = Index
Select Case Index
       Case 0
Forma2.Show
       Case 1
Formr445.Show
       Case 2
''Formj12.Show  订单审核
Formr449.Show
       Case 3
Forma132.Show
       Case 4
Formh101.Show
       Case 5
Forma111.Show
       Case 6
Forma9.Show
       Case 7
Forma1111.Show
       Case 8
Forma6.Show
       Case 9
Forma11.Show
       Case 10
Formy49.Show
       Case 11
Forma172.Show
       Case 12
Forma501.Show
       Case 13
Formc38.Show
       Case 14
Forma23.Show
       Case 15
Formd49.Show
       Case 16
Forma17.Show
       Case 17
Formc71.Show
       Case 18
Formj2.Show
       Case 19
Formj81.Show
       Case 20
Formj7.Show
       Case 21
Formh1.Show
       Case 22
Formh223.Show
       Case 23
Formh221.Show
       Case 24
Formh123.Show
       Case 25
Formh2.Show
       Case 26
Formh222.Show
       Case 27
Formh224.Show
       Case 28
Formd331.Show
       Case 29
Formd1112.Show
       Case 30
Formd2.Show
       Case 31
Formd45.Show
       Case 32
Formd47.Show
       Case 33
Forms501.Show
       Case 34
Forms504.Show
       Case 35
Forms500.Show
       Case 36
Forms507.Show
       Case 37
Forms502.Show
       Case 38
Forms503.Show
       Case 39
Forms506.Show
       Case 40
Forms508.Show
       Case 41
Forms54.Show
       Case 42
Forms51.Show
       Case 43
Forms509.Show
       Case 44
FormJ8.Show
       Case 45
Formc34.Show
       Case 46
Formc15.Show
       Case 47
FormC26.Show
       Case 48
Formc77.Show
       Case 49
Formc23.Show
       Case 50
Formc344.Show
       Case 51
'Formc15.Show
Formr447.Show
       Case 52
Forms3.Show
       Case 53
formr450.Show
       Case 54
Formr337.Show
       Case 55
Formr338.Show
       Case 56
Formr332.Show
       Case 57
Formc346.Show
       Case 58
'Formr334.Show
       Case 59
'Formr28.Show
       Case 60
Formr441.Show
       Case 61
Formr331.Show
       Case 62
'Formr50.Show
Forma55.Show
       Case 63
'''''''''''''''''原染色产量 Forms498.Show
Formd48.Show
       Case 64
Formh100.Show
       Case 65
Formr29.Show
       Case 66
Formr27.Show
       Case 67
Formy1.Show
       Case 68
Formy121.Show
       Case 69
Formy133.Show
       Case 70
Formy148.Show
       Case 71
Formy150.Show
       Case 72
Formy99.Show
       Case 73
Formy81.Show
       Case 74
Formy31.Show
       Case 75
Formy233.Show
       Case 76
Formy146.Show
       Case 77
Formy46.Show
       Case 78
Formy166.Show
       Case 79
Formw41.Show
       Case 80
Formw1127.Show
       Case 81
Formw1129.Show
       Case 82
Formw111.Show
       Case 83
Formw74.Show
       Case 84
Formw78.Show
       Case 85
Formc18.Show
      Case 86
Formc19.Show
       Case 87
Formc21.Show
       Case 88
Formc20.Show
       Case 89
FormW378.Show
       Case 90
Formw66.Show
       Case 91
Formc38.Show
       Case 93
Formw555.Show
       Case 94
Formw55.Show
       Case 95
Formw71.Show
       Case 96
Formw75.Show
       Case 97
Formw8.Show
       Case 98
Formw121.Show
       Case 99
Formc24.Show
       Case 100
Formc348.Show
       Case 101
Formr448.Show
       Case 102
Forma103.Show
       Case 103
Formy8.Show
       Case 104
Formj31.Show
       Case 105
Forma106.Show
       Case 106
Formw160.Show
       Case 107
Formw211.Show
       Case 108
Formj21.Show
       Case 109
Formw377.Show
       Case 110
Formw122.Show
       Case 111
Formw732.Show
       Case 112
Formw80.Show
       Case 113
Formw119.Show
       Case 114
Formm3.Show
       Case 115
Formw395.Show
       Case 116
Formy167.Show
       Case 117
Forms510.Show
       Case 118
Forms511.Show
       Case 119
Formh75.Show
       Case 120
Formr86.Show
       Case 121
Forma17.Show
       Case 122
Forma1.Show
       Case 123
Formj3.Show
       Case 124
Formj11.Show
'Forma173.Show
       Case 125
Formj6.Show
       Case 126
Formy82.Show
       Case 127
Formy83.Show
       Case 128
Formj17.Show
       Case 129
'Formy168.Show
       Case 131
Forms518.Show
      Case 130
Forms517.Show
      Case 132
Formh7.Show
      Case 133
Formr443.Show
      Case 134
Formj81.Show
      Case 135
Formc341.Show
      Case 136
Formc345.Show
      Case 137
Forma92.Show
     Case 138
Forma91.Show
     Case 139
Forma93.Show
     Case 140
Formy32.Show
     Case 141
Forma96.Show
     Case 142
Forma97.Show
     Case 143
Formj19.Show
     Case 144
Forma94.Show
     Case 145
Forma904.Show
     Case 146
Forms497.Show
     Case 147
Forma908.Show
     Case 148
Forma906.Show
     Case 149
Forma903.Show
     Case 150
Forma902.Show
     Case 151
Forma901.Show
     Case 152
Formh73.Show
     Case 153
Formw129.Show
    Case 154
Formj20.Show
     Case 155
Formd333.Show
     Case 156
Formy50.Show
     Case 157
Formy154.Show
     Case 158
Formr310.Show
     Case 159
Formy11.Show
     Case 160
Formw398.Show
     Case 161
Formc145.Show
     Case 162
Formy80.Show
     Case 163
Formj23.Show
     Case 164
Formr311.Show
     Case 165
Formr444.Show
     Case 166
Formr446.Show
     Case 171
Formr328.Show
     Case 173
Formr327.Show
     Case 200
 Formj81.Show
 Case 201
 Formd50.Show
 Case 202
 Forma188.Show
 Case 203
 Forma504.Show
 Case 204
 Forma505.Show
 Case 205
 Forms520.Show
 Case 206
 Formd440.Show
  Case 207
 Formy134.Show
 Case 208
 Formy100.Show
End Select

End Sub
Private Sub zqcd(Index As Integer)
Select Case Index
       Case 0
Forma2.ZOrder 0
       Case 1
Forma103.ZOrder 0
       Case 2
Formj12.ZOrder 0
       Case 3
Forma132.ZOrder 0
       Case 4
Formh101.ZOrder 0
       Case 5
Forma111.ZOrder 0
       Case 6
Forma9.ZOrder 0
       Case 7
Forma1111.ZOrder 0
       Case 8
Forma6.ZOrder 0
       Case 9
Forma11.ZOrder 0
       Case 10
Formy49.ZOrder 0
       Case 11
Forma172.ZOrder 0
       Case 12
Forma501.ZOrder 0
       Case 13
Formc38.ZOrder 0
       Case 14
Forma23.ZOrder 0
       Case 15
Formy167.ZOrder 0
       Case 16
Forma17.ZOrder 0
       Case 17
Forma21.ZOrder 0
       Case 18
Formj2.ZOrder 0
       Case 19
Forma22.ZOrder 0
       Case 20
Formj7.ZOrder 0
       Case 21
Formh1.ZOrder 0
       Case 22
Formh223.ZOrder 0
       Case 23
Formh221.ZOrder 0
       Case 24
Formh123.ZOrder 0
       Case 25
Formh2.ZOrder 0
       Case 26
Formh222.ZOrder 0
       Case 27
Formh224.ZOrder 0
       Case 28
Formd331.ZOrder 0
       Case 29
Formd1112.ZOrder 0
       Case 30
Formd2.ZOrder 0
       Case 31
Formd45.ZOrder 0
       Case 32
Formd47.ZOrder 0
       Case 33
Forms501.ZOrder 0
       Case 34
Forms504.ZOrder 0
       Case 35
Forms500.ZOrder 0
       Case 36
Forms507.ZOrder 0
       Case 37
Forms502.ZOrder 0
       Case 38
Forms503.ZOrder 0
       Case 39
Forms506.ZOrder 0
       Case 40
Forms508.ZOrder 0
       Case 41
Forms54.ZOrder 0
       Case 42
Forms51.ZOrder 0
       Case 43
Forms509.ZOrder 0
       Case 44
FormJ8.ZOrder 0
       Case 45
Formc34.ZOrder 0
       Case 46
Formc39.ZOrder 0
       Case 47
Formc23.ZOrder 0
       Case 48
Formc77.ZOrder 0
       Case 49
Formc23.ZOrder 0
       Case 50
Formc344.ZOrder 0
       Case 51
Formc15.ZOrder 0
       Case 52
Forms3.ZOrder 0
       Case 53
Formy149.ZOrder 0
       Case 54
Formr337.ZOrder 0
       Case 55
Formr338.ZOrder 0
       Case 56
Formr332.ZOrder 0
       Case 57
Formc346.ZOrder 0
       Case 58
Formr334.ZOrder 0
       Case 59
Formr28.ZOrder 0
       Case 60
Formr442.ZOrder 0
       Case 61
Formr331.ZOrder 0
       Case 62
Formr50.ZOrder 0
       Case 63
Forms498.ZOrder 0
       Case 64
Formh100.ZOrder 0
       Case 65
Formr29.ZOrder 0
       Case 66
Formr27.ZOrder 0
       Case 67
Formy1.ZOrder 0
       Case 68
Formy121.ZOrder 0
       Case 69
Formy133.ZOrder 0
       Case 70
Formy148.ZOrder 0
       Case 71
Formy150.ZOrder 0
       Case 72
Formy99.ZOrder 0
       Case 73
Formy577.ZOrder 0
       Case 74
Formy31.ZOrder 0
       Case 75
Formy233.ZOrder 0
       Case 76
Formy146.ZOrder 0
       Case 77
Formy46.ZOrder 0
       Case 78
Formy166.ZOrder 0
       Case 79
Formw41.ZOrder 0
       Case 80
Formw1127.ZOrder 0
       Case 81
Formw1129.ZOrder 0
       Case 82
Formw111.ZOrder 0
       Case 83
Formw74.ZOrder 0
       Case 84
Formw78.ZOrder 0
       Case 85
Formc18.ZOrder 0
      Case 86
Formc19.ZOrder 0
       Case 87
Formc21.ZOrder 0
       Case 88
Formc20.ZOrder 0

       Case 89
FormW378.ZOrder 0
       Case 90
Formw66.ZOrder 0
       Case 91
Formw1132.ZOrder 0
       Case 93
Formw555.ZOrder 0
       Case 94
Formw55.ZOrder 0
       Case 95
Formw71.ZOrder 0
       Case 96
Formw75.ZOrder 0
       Case 97
Formw8.ZOrder 0
       Case 98
Formw121.ZOrder 0
       Case 99
Formc24.ZOrder 0
       Case 100
Formc25.ZOrder 0
       Case 101
Formy81.ZOrder 0
       Case 102
Formj14.ZOrder 0
       Case 103
Formy8.ZOrder 0
       Case 104
Formj15.ZOrder 0
       Case 105
Forma106.ZOrder 0
       Case 106
Formw160.ZOrder 0
       Case 107
Formw211.ZOrder 0
       Case 108
Formj21.ZOrder 0
       Case 109
Formw377.ZOrder 0
       Case 110
Formw122.ZOrder 0
       Case 111
Formw732.ZOrder 0
       Case 112
Formw80.ZOrder 0
       Case 113
Formw119.ZOrder 0
       Case 114
Formm3.ZOrder 0
       Case 115
Formw395.ZOrder 0
       Case 116
Formy167.ZOrder 0
       Case 117
Forms510.ZOrder 0
       Case 118
Forms511.Show
       Case 119
Formh75.Show
       Case 120
Formr86.Show
       Case 121
Formj1.ZOrder 0
       Case 122
Formj16.ZOrder 0
       Case 123
Formj3.ZOrder 0
       Case 124
Formj11.ZOrder 0
'Forma173.Show
       Case 125
Formj6.ZOrder 0
       Case 126
Formy82.ZOrder 0
       Case 127
Formy83.ZOrder 0
       Case 128
Formj17.ZOrder 0
       Case 129
Formy168.ZOrder 0
       Case 131
Forms518.ZOrder 0
      Case 130
Forms517.ZOrder 0
      Case 132
Formh7.ZOrder 0
      Case 133
Forma27.ZOrder 0
      Case 134
Formj81.ZOrder 0
      Case 135
Formc341.ZOrder 0
      Case 136
Formc345.ZOrder 0
      Case 137
Forma92.ZOrder 0
     Case 138
Forma91.ZOrder 0
     Case 139
Forma93.ZOrder 0
     Case 140
Formy32.ZOrder 0
     Case 141
Forma96.ZOrder 0
     Case 142
Forma97.ZOrder 0
     Case 143
Formj19.ZOrder 0
     Case 144
Forma94.ZOrder 0
     Case 145
Forma904.ZOrder 0
     Case 146
Forms497.ZOrder 0
     Case 147
Forma908.ZOrder 0
     Case 148
Forma906.ZOrder 0
     Case 149
Forma903.ZOrder 0
     Case 150
Forma902.ZOrder 0
     Case 151
Forma901.ZOrder 0
     Case 152
Formh73.ZOrder 0
     Case 153
Formw129.ZOrder 0
    Case 154
Formj20.ZOrder 0
     Case 155
Formd333.ZOrder 0
     Case 156
Formy153.ZOrder 0
     Case 157
Formy154.ZOrder 0
     Case 158
Forma26.ZOrder 0
     Case 159
Formy11.ZOrder 0
     Case 160
Formw398.ZOrder 0
     Case 161
Formc145.ZOrder 0
     Case 162
Formy80.ZOrder 0
     Case 163
Formj23.ZOrder 0
End Select
End Sub


Private Sub Label2_Click()
If InStr(yhm, "cw") > 0 Or InStr(yhm, "root") > 0 Or InStr(yhm, "CW") > 0 Or InStr(yhm, "ROOT") > 0 Then
xtxxjm = "停止"
End If
End Sub

Private Sub TCXT_Click()
End
End Sub

