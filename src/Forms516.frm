VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forms516 
   BackColor       =   &H00C0E0FF&
   Caption         =   "验布记录"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   13995
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "Text2"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   125
      Text            =   "Text2"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "织布疵点(处)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   840
      TabIndex        =   78
      Top             =   1080
      Width           =   7335
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   600
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1200
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1800
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2400
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3000
         TabIndex        =   85
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   3600
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   4200
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   4800
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   5400
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   6000
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   11
         Left            =   6600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Text            =   "Forms516.frx":0000
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "长残"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   124
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "割车"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   720
         TabIndex        =   123
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "破洞"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   1200
         TabIndex        =   122
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "掉扣"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   1800
         TabIndex        =   121
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "希路针"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   2400
         TabIndex        =   120
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "横道"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   3000
         TabIndex        =   119
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "油污"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   3600
         TabIndex        =   118
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "花毛织入"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   4200
         TabIndex        =   117
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "花针"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   8
         Left            =   4800
         TabIndex        =   116
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "断氨纶丝"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   9
         Left            =   5400
         TabIndex        =   115
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   114
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   113
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   112
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   111
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   110
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3000
         TabIndex        =   109
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3600
         TabIndex        =   108
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4200
         TabIndex        =   107
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4800
         TabIndex        =   106
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   105
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   104
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   103
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   102
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   101
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   100
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3000
         TabIndex        =   99
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3600
         TabIndex        =   98
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4200
         TabIndex        =   97
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4800
         TabIndex        =   96
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   5400
         TabIndex        =   95
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6000
         TabIndex        =   94
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6000
         TabIndex        =   93
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "反丝"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   10
         Left            =   6000
         TabIndex        =   92
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   29
         Left            =   6600
         TabIndex        =   91
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "纱线疵点(处)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   1
      Left            =   8400
      TabIndex        =   69
      Top             =   1080
      Width           =   3375
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   240
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   16
         Left            =   1080
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "粗细纱"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   22
         Left            =   240
         TabIndex        =   77
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   76
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   75
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "异色纤维"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   37
         Left            =   1080
         TabIndex        =   74
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   1080
         TabIndex        =   73
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   1080
         TabIndex        =   72
         Top             =   2760
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "疵点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   2
      Left            =   12120
      TabIndex        =   66
      Top             =   1080
      Width           =   1095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   120
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "疵点合计"
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
         Index           =   11
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "织布疵点(米)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   0
      Left            =   840
      TabIndex        =   23
      Top             =   4440
      Width           =   6735
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   840
         TabIndex        =   32
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   1440
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   4
         Left            =   2040
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   5
         Left            =   2640
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   6
         Left            =   3240
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   8
         Left            =   3960
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   9
         Left            =   4560
         TabIndex        =   26
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   10
         Left            =   5160
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   11
         Left            =   5760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "Forms516.frx":0006
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "断氨纶丝"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   14
         Left            =   4560
         TabIndex        =   61
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "花针"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   16
         Left            =   3960
         TabIndex        =   60
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "油污"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   24
         Left            =   3240
         TabIndex        =   59
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "横道"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   25
         Left            =   2640
         TabIndex        =   58
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "希路针"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   26
         Left            =   2040
         TabIndex        =   57
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "掉扣"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   27
         Left            =   1440
         TabIndex        =   56
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "破洞"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   28
         Left            =   840
         TabIndex        =   55
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "长残"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   30
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   51
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   50
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2040
         TabIndex        =   49
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   48
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3240
         TabIndex        =   47
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3960
         TabIndex        =   46
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   45
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   44
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   43
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   42
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   41
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3240
         TabIndex        =   40
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3960
         TabIndex        =   39
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4560
         TabIndex        =   38
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "反丝"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   18
         Left            =   5160
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   36
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   35
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   33
         Left            =   5760
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "纱线疵点(米)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   1
      Left            =   8400
      TabIndex        =   14
      Top             =   4440
      Width           =   3375
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   12
         Left            =   240
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   16
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "粗细纱"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   32
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   21
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "异色纤维"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   38
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   1080
         TabIndex        =   18
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   1080
         TabIndex        =   17
         Top             =   3000
         Width           =   495
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Forms516.frx":000C
      Left            =   11640
      List            =   "Forms516.frx":0019
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "疵点"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   2
      Left            =   12120
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   120
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "疵点合计"
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
         Index           =   36
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   3480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   4200
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   5280
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   3000
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   2040
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   480
      Width           =   1335
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms516.frx":002F
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   8160
      Width           =   12375
      _cx             =   21828
      _cy             =   2143
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2160
      Top             =   9480
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "匹号"
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
      Index           =   17
      Left            =   3240
      TabIndex        =   158
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "织号"
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
      Index           =   15
      Left            =   840
      TabIndex        =   157
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "等级"
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
      Index           =   12
      Left            =   10800
      TabIndex        =   156
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "其它"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   13
      Left            =   1560
      TabIndex        =   155
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   960
      TabIndex        =   154
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   600
      TabIndex        =   153
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "杂纤维"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   23
      Left            =   1680
      TabIndex        =   152
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   151
      Top             =   8760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1680
      TabIndex        =   150
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "杂纤维"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   31
      Left            =   1800
      TabIndex        =   149
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1560
      TabIndex        =   148
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1320
      TabIndex        =   147
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "大肚纱"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   21
      Left            =   1680
      TabIndex        =   146
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   720
      TabIndex        =   145
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   1080
      TabIndex        =   144
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1800
      TabIndex        =   143
      Top             =   9600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1440
      TabIndex        =   142
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "死棉结"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   20
      Left            =   1440
      TabIndex        =   141
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1560
      TabIndex        =   140
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   960
      TabIndex        =   139
      Top             =   10320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "死棉结"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   34
      Left            =   1560
      TabIndex        =   138
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   6240
      TabIndex        =   137
      Top             =   9120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   1560
      TabIndex        =   136
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "横道"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   19
      Left            =   1440
      TabIndex        =   135
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   1560
      TabIndex        =   134
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   960
      TabIndex        =   133
      Top             =   10320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "横道"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   35
      Left            =   1560
      TabIndex        =   132
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   131
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   130
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   129
      Top             =   10440
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1440
      TabIndex        =   128
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "挡车"
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
      Left            =   4560
      TabIndex        =   127
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Forms516"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text4 = "" Then
MsgBox ("请输入挡车编号")
Exit Sub
End If

If Combo1 = "" Then
MsgBox ("请输入等级")
Exit Sub
End If

Adodc1.Recordset.AddNew
For i = 0 To 1
Adodc1.Recordset.Fields(i) = Text2(i)
Next

For i = 0 To 17
Adodc1.Recordset.Fields(2 + i) = Text1(i)
Next

For i = 0 To 17
Adodc1.Recordset.Fields(20 + i) = Text3(i)
Next

Adodc1.Recordset.Fields(38) = Combo1
Adodc1.Recordset.Fields(39) = Text4
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text4 = "" Then
MsgBox ("请输入挡车编号")
Exit Sub
End If


If Combo1 = "" Then
MsgBox ("请输入等级")
Exit Sub
End If

For i = 0 To 1
Adodc1.Recordset.Fields(i) = Text2(i)
Next

For i = 0 To 17
Adodc1.Recordset.Fields(2 + i) = Text1(i)
Next

For i = 0 To 17
Adodc1.Recordset.Fields(20 + i) = Text3(i)
Next

Adodc1.Recordset.Fields(38) = Combo1
Adodc1.Recordset.Fields(39) = Text4
Adodc1.Recordset.Update
Adodc1.Refresh

End Sub

Private Sub Command4_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh
End Sub

Private Sub Form_Load()
For i = 0 To 17
Text1(i) = ""
Text3(i) = ""
Next
Combo1 = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from zbzjbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
Adodc1.Refresh
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case Index
Text1(Index) = Val(Text1(Index)) + 1
End Select
End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case Index
If Val(Text1(Index)) <= 0 Then Exit Sub
Text1(Index) = Val(Text1(Index)) - 1
End Select
End Sub

Private Sub Label4_Click(Index As Integer)
Select Case Index
       Case Index
Text3(Index) = Val(Text3(Index)) + 1
End Select
End Sub

Private Sub Label5_Click(Index As Integer)
Select Case Index
       Case Index
If Val(Text3(Index)) <= 0 Then Exit Sub
Text3(Index) = Val(Text3(Index)) - 1
End Select
End Sub

Private Sub Label6_Click()
'If Forms508.Option3.Value = True Then
'Text4 = Forms508.Text2(10)
'End If
End Sub

Private Sub Label6_dblClick()
'If Forms508.Option3.Value = True Then
'Text4 = Forms508.Text2(4)
'End If
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case Index
If Index < 11 Then
bh = 0
For i = 0 To 10
bh = bh + Val(Text1(i))
Next
End If
Text1(17) = bh
End Select
End Sub

Private Sub Text2_Change(Index As Integer)
Select Case Index
       Case 1
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from zbzjbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
Adodc1.Refresh
End Select
End Sub

Private Sub Text3_Change(Index As Integer)
Select Case Index
       Case Index
If Index < 11 Then
bh = 0
For i = 0 To 10
bh = bh + Val(Text3(i))
Next
End If
Text3(17) = bh
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
For i = 0 To 16
Text1(i) = Adodc1.Recordset.Fields(i + 2)
Next
For i = 0 To 16
Text3(i) = Adodc1.Recordset.Fields(i + 20)
Next

Combo1 = Adodc1.Recordset.Fields(38)
Text4 = Adodc1.Recordset.Fields(39)
End Sub

