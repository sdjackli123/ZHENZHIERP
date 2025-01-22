VERSION 5.00
Begin VB.Form Formm2 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15240
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "用户管理"
      DownPicture     =   "Formm2.frx":0000
      Height          =   375
      Left            =   960
      MaskColor       =   &H00C0C0FF&
      MouseIcon       =   "Formm2.frx":014A
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "用户切换"
      DownPicture     =   "Formm2.frx":0294
      Height          =   375
      Left            =   1920
      MaskColor       =   &H00C0C0FF&
      MouseIcon       =   "Formm2.frx":03DE
      MousePointer    =   4  'Icon
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "染缸排缸"
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
      Left            =   10560
      MouseIcon       =   "Formm2.frx":0528
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   480
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯配缸"
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
      Left            =   8160
      MouseIcon       =   "Formm2.frx":16BBA
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   0
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "配缸检测"
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
      Left            =   13200
      MouseIcon       =   "Formm2.frx":2D24C
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   360
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分缸查询"
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
      Left            =   13200
      MouseIcon       =   "Formm2.frx":438DE
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "申请确认"
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
      Index           =   172
      Left            =   3480
      MouseIcon       =   "Formm2.frx":59F70
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   120
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库台账"
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
      Left            =   360
      MouseIcon       =   "Formm2.frx":70602
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   1080
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "串口设置"
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
      Left            =   8760
      MouseIcon       =   "Formm2.frx":86C94
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   1200
      Width           =   1740
   End
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
      Index           =   13
      Left            =   4800
      MouseIcon       =   "Formm2.frx":9D326
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "染料台账"
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
      Index           =   108
      Left            =   10560
      MouseIcon       =   "Formm2.frx":B39B8
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "计划预警设置"
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
      Left            =   4560
      MouseIcon       =   "Formm2.frx":CA04A
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   4200
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "原液预警设置"
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
      Left            =   600
      MouseIcon       =   "Formm2.frx":E06DC
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   4320
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "原液预警信息"
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
      Index           =   157
      Left            =   8040
      MouseIcon       =   "Formm2.frx":F6D6E
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   4320
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "机号工艺"
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
      Left            =   960
      MouseIcon       =   "Formm2.frx":10D400
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   3120
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
      Index           =   9
      Left            =   720
      MouseIcon       =   "Formm2.frx":123A92
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "备活染料"
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
      Index           =   146
      Left            =   11880
      MouseIcon       =   "Formm2.frx":13A124
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   2280
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   540
      Index           =   8
      Left            =   12600
      Picture         =   "Formm2.frx":1507B6
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   345
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
      Index           =   4
      Left            =   11040
      MouseIcon       =   "Formm2.frx":1509CA
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   7080
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   540
      Index           =   19
      Left            =   12000
      Picture         =   "Formm2.frx":16705C
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   345
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
      Index           =   2
      Left            =   11280
      MouseIcon       =   "Formm2.frx":167270
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   6240
      Width           =   1740
   End
   Begin VB.Image Image6 
      Height          =   495
      Index           =   0
      Left            =   12000
      Picture         =   "Formm2.frx":17D902
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯盘存"
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
      Left            =   11280
      MouseIcon       =   "Formm2.frx":17DB16
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   9000
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
      Index           =   127
      Left            =   11280
      MouseIcon       =   "Formm2.frx":1941A8
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   7680
      Width           =   1740
   End
   Begin VB.Image Image6 
      Height          =   615
      Index           =   2
      Left            =   12000
      Picture         =   "Formm2.frx":1AA83A
      Stretch         =   -1  'True
      Top             =   6840
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
      Index           =   44
      Left            =   4080
      MouseIcon       =   "Formm2.frx":1AAA4E
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   2760
      Width           =   1740
   End
   Begin VB.Image Image6 
      Height          =   495
      Index           =   3
      Left            =   4680
      Picture         =   "Formm2.frx":1C10E0
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "光坯盘存"
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
      Index           =   128
      Left            =   3960
      MouseIcon       =   "Formm2.frx":1C12F4
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   7200
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
      Index           =   129
      Left            =   3960
      MouseIcon       =   "Formm2.frx":1D7986
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   5160
      Width           =   1740
   End
   Begin VB.Image Image6 
      Height          =   495
      Index           =   4
      Left            =   4560
      Picture         =   "Formm2.frx":1EE018
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "光坯台账"
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
      Left            =   4080
      MouseIcon       =   "Formm2.frx":1EE22C
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   6480
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯台账"
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
      Left            =   4440
      MouseIcon       =   "Formm2.frx":2048BE
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   5760
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "复核信息"
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
      Left            =   3240
      MouseIcon       =   "Formm2.frx":21AF50
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   1680
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询设置"
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
      Left            =   11880
      MouseIcon       =   "Formm2.frx":2315E2
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   5280
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "年初设置"
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
      Left            =   11880
      MouseIcon       =   "Formm2.frx":247C74
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   3120
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "负债设置"
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
      Left            =   11880
      MouseIcon       =   "Formm2.frx":25E306
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   3840
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "利润设置"
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
      Left            =   11880
      MouseIcon       =   "Formm2.frx":274998
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   4560
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "报价转入"
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
      Index           =   116
      Left            =   9720
      MouseIcon       =   "Formm2.frx":28B02A
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   1920
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "凭证填制"
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
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   1200
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "凭证审核"
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
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "账本操作"
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
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   3600
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "账本查询"
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
      Index           =   102
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6000
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "会计报表"
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
      Index           =   104
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4800
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "账期查询"
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
      Index           =   105
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   7200
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   4
      Left            =   7560
      Picture         =   "Formm2.frx":2A16BC
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   5
      Left            =   7560
      Picture         =   "Formm2.frx":2A18D0
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   6
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   14
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   7
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "发生查询"
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
      Left            =   6960
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   8400
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   23
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "材料修正"
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
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库存查询"
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
      Index           =   160
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分出库查询"
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
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分入库查询"
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
      Left            =   1200
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分材料出库"
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
      Left            =   9240
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   405
      Index           =   1
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   690
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   25
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   26
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分盘存操作"
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
      Left            =   9240
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库存记录"
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
      Index           =   159
      Left            =   9240
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分材料盘存"
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
      Left            =   9240
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   27
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "传票查询"
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
      Left            =   2160
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3600
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "特种记账"
      Height          =   375
      Index           =   62
      Left            =   5400
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "应付凭证"
      Height          =   375
      Index           =   20
      Left            =   3960
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "凭证生成"
      Height          =   495
      Index           =   32
      Left            =   5880
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "凭证生成"
      Height          =   495
      Index           =   31
      Left            =   7200
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15210
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   840
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "费用记录"
      Height          =   495
      Index           =   26
      Left            =   13680
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "加工费用查询、凭证生成"
      Height          =   495
      Index           =   38
      Left            =   13680
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "凭证审核"
      Height          =   495
      Index           =   40
      Left            =   13680
      TabIndex        =   4
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "账本预览"
      Height          =   495
      Index           =   42
      Left            =   13680
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "染色单价"
      Height          =   375
      Index           =   88
      Left            =   13560
      TabIndex        =   2
      Top             =   8520
      Width           =   1335
   End
End
Attribute VB_Name = "Formm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
Formy1.Show
       Case 1
Formy306.Show
       Case 2
Formy19.Show
       Case 3
Forml503.Show
       Case 4
Formy4.Show
       Case 5
Formy30.Show
       Case 6
Formy302.Show
       Case 7
Formy301.Show
       Case 8
Formy303.Show
       Case 9
Formc46.Show
       Case 10
Forml602.Show
       Case 11
Forml605.Show
       Case 12
Forms203.Show
       Case 13
Forml191.Show
       Case 14
Formb17.Show
       Case 15
Forml603.Show
       Case 16
Forml604.Show
       Case 17
Formy195.Show
       Case 18
Formw211.Show
       Case 19
Formw31.Show
       Case 20
Formw31.Show
       Case 21

       Case 22
Formw91.Show
       Case 23
Formy203.Show
       Case 24
Formw92.Show
       Case 25
Formw94.Show
       Case 27
Formb16.Show
       Case 28
Formw2.Show
       Case 29
Formw114.Show
       Case 30
Formw2.Show

       Case 33
Formy306.Show
       Case 34
Formc47.Show
       Case 35
Formw10.Show
       Case 36
Formw218.Show
       Case 37
Formw205.Show
       Case 39
Formc22.Show
       Case 41
Formw35.Show
       Case 43
Formw117.Show
       Case 44
Formy305.Show
       Case 45
Formy25.Show
       Case 46
Formy45.Show
       Case 47
Formw95.Show
       Case 48
Formw93.Show
       Case 49
Formw97.Show
       Case 50
Formw203.Show
       Case 51
Formw555.Show
       Case 52
Formw55.Show
       Case 53
Formw24.Show
       Case 54
Formw37.Show
       Case 55
Formw36.Show
       Case 56
Formw1.Show
       Case 57
Formw111.Show
       Case 58
Formw10.Show
       Case 59
Formw112.Show
       Case 60
Formw113.Show
       Case 61
Formw29.Show
       Case 62
Formw114.Show
       Case 63
Formw116.Show
       Case 64
Formw115.Show
       Case 65
Formw66.Show
       Case 66
Formw81.Show
       Case 67
Formw8.Show
       Case 68
Formw82.Show
       Case 69
Formw31.Show
       Case 70
Formw2.Show
       Case 71
Formw206.Show
       Case 72
Formw211.Show
       Case 73
Formw208.Show
       Case 74
Formw212.Show
       Case 75
Formw217.Show
       Case 76
Formw216.Show
       Case 77
Formw27.Show
       Case 78
Formw32.Show
       Case 79
Formw28.Show
       Case 80
Formw35.Show
       Case 81
Formw337.Show
       Case 82
Formw332.Show
       Case 83
Formw338.Show
       Case 84
Formw48.Show
       Case 85
Formw40.Show
      Case 86
Formw49.Show
       Case 87
Formc21.Show
       Case 89
Formc34.Show
       Case 90
Formc2.Show
       Case 91
Formc45.Show
End Select

End Sub

Private Sub TCXT_Click()
End
End Sub

