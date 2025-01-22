VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Formr339 
   BackColor       =   &H00C0E0FF&
   Caption         =   "粉体称量"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8280
      Top             =   840
   End
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
      Tab             =   1
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
      TabPicture(0)   =   "Formr339.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "配料信息"
      TabPicture(1)   =   "Formr339.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   10215
         Index           =   0
         Left            =   0
         ScaleHeight     =   10155
         ScaleWidth      =   18435
         TabIndex        =   29
         Top             =   600
         Width           =   18495
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "通讯口操作："
            Height          =   1335
            Index           =   0
            Left            =   8160
            TabIndex        =   118
            Top             =   240
            Width           =   6135
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   2880
               TabIndex        =   122
               Text            =   "Text6"
               Top             =   840
               Width           =   1575
            End
            Begin VB.CommandButton Command11 
               BackColor       =   &H00C0C0FF&
               Caption         =   "关闭串口"
               Height          =   375
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   121
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H00C0C0FF&
               Caption         =   "打开串口"
               Height          =   375
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   240
               Width           =   1095
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               ItemData        =   "Formr339.frx":0038
               Left            =   240
               List            =   "Formr339.frx":003A
               TabIndex        =   119
               Text            =   "COM1"
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "请打开串口"
               ForeColor       =   &H00000040&
               Height          =   180
               Index           =   4
               Left            =   240
               TabIndex        =   125
               Top             =   945
               Width           =   900
            End
            Begin VB.Label Label244 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "通讯状态："
               ForeColor       =   &H00000040&
               Height          =   300
               Index           =   1
               Left            =   1920
               TabIndex        =   124
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "端口号："
               Height          =   180
               Left            =   240
               TabIndex        =   123
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "实时读Y0--Y7"
            Height          =   3135
            Index           =   0
            Left            =   10200
            TabIndex        =   71
            Top             =   1560
            Visible         =   0   'False
            Width           =   6615
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   8
               Left            =   480
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
               Top             =   480
               Width           =   255
            End
            Begin VB.Timer Timer4 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   6000
               Top             =   360
            End
            Begin VB.Timer Timer3 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   6000
               Top             =   840
            End
            Begin VB.PictureBox Picture5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   3120
               ScaleHeight     =   255
               ScaleWidth      =   255
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
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
               TabIndex        =   80
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
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
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
               Top             =   1320
               Width           =   255
            End
            Begin MSCommLib.MSComm MSComm4 
               Left            =   5880
               Top             =   1320
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
               TabIndex        =   117
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
               TabIndex        =   116
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
               TabIndex        =   115
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
               TabIndex        =   114
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
               TabIndex        =   113
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
               TabIndex        =   112
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
               TabIndex        =   111
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
               TabIndex        =   110
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
               TabIndex        =   109
               Top             =   240
               Width           =   210
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
               TabIndex        =   108
               Top             =   240
               Width           =   210
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
               TabIndex        =   104
               Top             =   240
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
               TabIndex        =   103
               Top             =   240
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
               TabIndex        =   102
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
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
               Top             =   1080
               Width           =   315
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
               TabIndex        =   96
               Top             =   1080
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "数值操作"
            Height          =   255
            Left            =   10200
            TabIndex        =   52
            Top             =   2400
            Visible         =   0   'False
            Width           =   6615
            Begin VB.TextBox Text7 
               Height          =   390
               Left            =   3840
               TabIndex        =   66
               Text            =   "Text7"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   960
               TabIndex        =   65
               Text            =   "Text5"
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton Command9 
               Caption         =   "写入"
               Height          =   420
               Left            =   4440
               TabIndex        =   64
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton Command8 
               Caption         =   "读值"
               Height          =   420
               Left            =   3840
               TabIndex        =   63
               Top             =   960
               Width           =   615
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
               Text            =   "text8"
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   240
               TabIndex        =   58
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton Option13 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   61
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option12 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   60
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "D"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00C0FFC0&
               Caption         =   "位数"
               Height          =   615
               Left            =   240
               TabIndex        =   54
               Top             =   960
               Width           =   2895
               Begin VB.OptionButton Option14 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "浮点"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   57
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option5 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "32位"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   56
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "16位"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   960
               TabIndex        =   53
               Text            =   "Text10"
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "显示读出数值："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   0
               Left            =   2280
               TabIndex        =   70
               Top             =   2160
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
               TabIndex        =   69
               Top             =   1800
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
               TabIndex        =   68
               Top             =   1800
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "读地址："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   51
               Left            =   360
               TabIndex        =   67
               Top             =   2160
               Width           =   720
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "状态操作"
            Height          =   135
            Left            =   10200
            TabIndex        =   38
            Top             =   2640
            Visible         =   0   'False
            Width           =   6615
            Begin VB.CommandButton Command7 
               Caption         =   "查询当前状态"
               Height          =   420
               Left            =   3120
               TabIndex        =   49
               Top             =   1560
               Width           =   1335
            End
            Begin VB.CommandButton Command6 
               Caption         =   "置位"
               Height          =   420
               Left            =   360
               TabIndex        =   48
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "复位"
               Height          =   420
               Left            =   1800
               TabIndex        =   47
               Top             =   1560
               Width           =   1215
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
               TabIndex        =   46
               Text            =   "text4"
               Top             =   960
               Width           =   1575
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   345
               TabIndex        =   39
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton Option10 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "S"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   45
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option8 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   44
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option9 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   43
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option7 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "M"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   42
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Option6 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Y"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   41
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option11 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "X"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   240
                  Width           =   495
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "地址："
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   36
               Left            =   240
               TabIndex        =   51
               Top             =   1080
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "状态指示灯"
               ForeColor       =   &H000040C0&
               Height          =   180
               Index           =   67
               Left            =   2400
               TabIndex        =   50
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
            Caption         =   "配料信息"
            Height          =   975
            Left            =   3120
            TabIndex        =   35
            Top             =   600
            Width           =   3255
            Begin VB.OptionButton Option2 
               BackColor       =   &H0000FF00&
               Caption         =   "已称量"
               Height          =   495
               Left            =   1680
               TabIndex        =   37
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H000000FF&
               Caption         =   "未称量"
               Height          =   495
               Left            =   240
               TabIndex        =   36
               Top             =   240
               Width           =   1215
            End
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
            TabIndex        =   34
            Top             =   720
            Width           =   1455
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
            TabIndex        =   33
            Text            =   "Text9"
            Top             =   2040
            Width           =   3735
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   0
            Left            =   6720
            TabIndex        =   32
            Text            =   "Text12"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   1
            Left            =   7560
            TabIndex        =   31
            Text            =   "Text12"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   375
            Index           =   2
            Left            =   8400
            TabIndex        =   30
            Text            =   "Text12"
            Top             =   1680
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1680
            TabIndex        =   126
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   255
            Format          =   329318401
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   127
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   329318401
            CurrentDate     =   36892
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
            TabIndex        =   172
            Top             =   2040
            Width           =   1815
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
            TabIndex        =   171
            Top             =   600
            Width           =   1335
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
            TabIndex        =   170
            Top             =   1200
            Width           =   1335
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
            TabIndex        =   169
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
            Index           =   28
            Left            =   7680
            TabIndex        =   168
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
            Index           =   27
            Left            =   7680
            TabIndex        =   167
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
            Index           =   26
            Left            =   7680
            TabIndex        =   166
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
            Index           =   25
            Left            =   7680
            TabIndex        =   165
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
            Index           =   24
            Left            =   7680
            TabIndex        =   164
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
            Index           =   23
            Left            =   5880
            TabIndex        =   163
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
            Index           =   22
            Left            =   5880
            TabIndex        =   162
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
            Index           =   21
            Left            =   5880
            TabIndex        =   161
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
            Index           =   20
            Left            =   5880
            TabIndex        =   160
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
            Index           =   19
            Left            =   5880
            TabIndex        =   159
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
            Index           =   18
            Left            =   5880
            TabIndex        =   158
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
            Index           =   17
            Left            =   4080
            TabIndex        =   157
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
            Index           =   16
            Left            =   4080
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
            Index           =   15
            Left            =   4080
            TabIndex        =   155
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
            Index           =   14
            Left            =   4080
            TabIndex        =   154
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
            Index           =   13
            Left            =   4080
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
            Index           =   12
            Left            =   4080
            TabIndex        =   152
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
            Index           =   11
            Left            =   2280
            TabIndex        =   151
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
            Index           =   9
            Left            =   2280
            TabIndex        =   149
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
            Index           =   8
            Left            =   2280
            TabIndex        =   148
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
            Index           =   7
            Left            =   2280
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
            Index           =   6
            Left            =   2280
            TabIndex        =   146
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
            Index           =   5
            Left            =   480
            TabIndex        =   145
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
            Index           =   4
            Left            =   480
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
            Index           =   3
            Left            =   480
            TabIndex        =   143
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
            Index           =   2
            Left            =   480
            TabIndex        =   142
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
            Index           =   1
            Left            =   480
            TabIndex        =   141
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
            Height          =   855
            Index           =   0
            Left            =   480
            TabIndex        =   140
            Top             =   3480
            Width           =   1455
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
            TabIndex        =   139
            Top             =   2760
            Width           =   8655
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
            TabIndex        =   138
            Top             =   2760
            Width           =   3975
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
            TabIndex        =   137
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
            Height          =   975
            Index           =   8
            Left            =   11760
            TabIndex        =   136
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
            Index           =   7
            Left            =   11760
            TabIndex        =   135
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
            Index           =   6
            Left            =   11760
            TabIndex        =   134
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
            Height          =   855
            Index           =   5
            Left            =   11760
            TabIndex        =   133
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
            Index           =   4
            Left            =   9600
            TabIndex        =   132
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
            Height          =   975
            Index           =   3
            Left            =   9600
            TabIndex        =   131
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
            Index           =   2
            Left            =   9600
            TabIndex        =   130
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
            Index           =   1
            Left            =   9600
            TabIndex        =   129
            Top             =   4680
            Width           =   1815
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
            TabIndex        =   128
            Top             =   3480
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Index           =   1
         Left            =   -75000
         ScaleHeight     =   10275
         ScaleWidth      =   18315
         TabIndex        =   1
         Top             =   600
         Width           =   18375
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   2040
            TabIndex        =   11
            Text            =   "Text3"
            Top             =   1560
            Width           =   3855
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   2040
            TabIndex        =   10
            Text            =   "Text2"
            Top             =   600
            Width           =   1935
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
            TabIndex        =   9
            Top             =   1200
            Width           =   1455
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
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   2640
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
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   3600
            Width           =   1575
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
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   4440
            Width           =   5655
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
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   6600
            Width           =   5415
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   7080
            Top             =   240
         End
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   7680
            Top             =   240
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "关闭串口"
            Height          =   495
            Left            =   13560
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
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
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   600
            Width           =   1455
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
            TabIndex        =   2
            Text            =   "Text11"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Timer Timer6 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   8760
            Top             =   240
         End
         Begin VB.Timer Timer7 
            Enabled         =   0   'False
            Interval        =   500
            Left            =   9240
            Top             =   240
         End
         Begin VB.Timer Timer8 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   9840
            Top             =   240
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
            Bindings        =   "Formr339.frx":003C
            Height          =   5535
            Left            =   240
            TabIndex        =   12
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
            TabIndex        =   28
            Top             =   600
            Width           =   1935
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
            TabIndex        =   27
            Top             =   3600
            Visible         =   0   'False
            Width           =   4095
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
            TabIndex        =   26
            Top             =   1560
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
            TabIndex        =   25
            Top             =   600
            Width           =   1815
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
            TabIndex        =   24
            Top             =   2280
            Width           =   1815
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
            TabIndex        =   23
            Top             =   2640
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
            TabIndex        =   22
            Top             =   3600
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
            TabIndex        =   21
            Top             =   4440
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
            TabIndex        =   20
            Top             =   5880
            Width           =   1215
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
            TabIndex        =   19
            Top             =   600
            Width           =   1215
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
            TabIndex        =   18
            Top             =   600
            Width           =   4215
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
            TabIndex        =   17
            Top             =   5880
            Visible         =   0   'False
            Width           =   2175
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
            TabIndex        =   16
            Top             =   7440
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
            TabIndex        =   15
            Top             =   7440
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
            TabIndex        =   14
            Top             =   6600
            Width           =   615
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
            TabIndex        =   13
            Top             =   6600
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "Formr339"
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
Dim dzdq(3) As Double ''''电子称变量和判断染料编号
Dim dzbl(4) As Double ''''给电子称传输变量
Dim dzdqpd As Integer  ''''那个电子称
Dim bzzl As Double   '''包装重量
Dim htzl As Double ''''换筒变量
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   '浮点数处理
Dim MXH  As Integer    '''''''''循环读M

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
 b = CloseComm(MSComm1)
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

Label4.Caption = ""

Text3 = ""
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

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

Private Sub Label15_Click()
On Error Resume Next
       ReDim WriteData(0 To 14) As String  ''''''写入个数
       Dim DataW As String    '*浮点数的中间处理变量；
       Dim Data10(7) As Single   '*浮点数的中间处理变量；
       Dim Buffer(3) As Byte   '*浮点数的中间处理变量；
       
SBBH = 0   ''''''''''''''''''''        设备编号
dzbl(1) = 0
dzbl(2) = 0
dzbl(3) = 0
dzdqpd = 0
dzbl(4) = 0
 
       For i = 0 To 3
       Data10(i) = Val(dzbl(i + 1))
       CopyMemory Buffer(0), Data10(i), 4
       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
       Next
       a = gk528WriteDevice("D600", 8, WriteData())
 RWorder = 6
 RWcomm = True
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
       Case 6
Adodc5.RecordSource = "select 包装数量 from ftsb where 粉体名称='" & Text1(0) & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
Call Label15_Click
bzzl = Val(Adodc5.Recordset.Fields(0)) * 1000
Else
bzzl = 0
End If

     Case 8

cll = Format(Val(Text1(3)) / 1000, "#0.00000")   ''''''''''称量单位g转换成kg
sql1 = "UPDATE pldr SET 实际称量=(isnull(实际称量,0)+'" & cll & "'),称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
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
''''''''''''''''''''''''''''''''关闭盖子
'Text4 = 1

'''''''''''''''''''''''''''''''''''''''''
'If MSComm1.PortOpen = True Then
'MSComm1.Output = Chr$(27) + "t"
'End If

'If MSComm2.PortOpen = True Then
'MSComm2.Output = Chr$(27) + "t"
'End If
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
          Case 0   'read d704-708
   '
                         For i = 0 To 2
                         dataCl = Mid(PLCText, i * 8 + 7, 2) + Mid(PLCText, i * 8 + 5, 2) + Mid(PLCText, i * 8 + 3, 2) + Mid(PLCText, i * 8 + 1, 2)
                         Data10 = Val("&H" & dataCl) '*转换为十进制
                         CopyMemory Data, Data10, 4 '*转换为浮点数，调用模块中的COPY声明,意义为:将L复制给F1,位数为四位.
                         dzdq(i + 1) = Format(CStr(Data), "#0.000")
                         Next i
                         If dzdqpd = 1 Then
                         Text1(3) = Format(Val(dzdq(1)) + bzzl, "#0.00")
                         End If
                         
                         If dzdqpd = 2 Then
                         Text1(3) = Format(Val(dzdq(2)) + bzzl, "#0")
                         End If
                         
                         Text12(0) = Val(dzdq(1))
                         Text12(1) = Val(dzdq(2))
                         Text12(2) = Val(dzdq(3))
               
          Case 6, 7, 8  '写 置，复位
               Order = 0
   End Select

   Timer3.Enabled = True

End Sub


Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 3
If Val(Text1(3)) > 0 And Val(Text1(2)) <= (Val(Text1(3)) + 0.02) And Val(Text1(2)) >= (Val(Text1(3)) - 0.02) And Val(Text1(2)) > 0 Then     '''''判断是否保存
Timer2.Enabled = True
ksjs = 0
End If
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
If Len(Text2) = 10 Then
Adodc7.RecordSource = "select 料单编号 from ldkh where 卡号编号='" & Text2 & "'"
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
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from FTBZXS where 料单编号='" & Text3 & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Formr333.Text3 = Text3
Formr333.Show
End If

Call VQJC

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
If Val(Text1(3)) > 0 And Val(Text1(2)) <= (Val(Text1(3)) + 0.05) And Val(Text1(2)) >= (Val(Text1(3)) - 0.05) And Val(Text1(2)) > 0 Then  ''''''误差在——+0.02g
ksjs = ksjs + 1
Beep 1000, 50
If ksjs / 2 = Int(ksjs / 2) Then
Text1(0).ForeColor = &HFF&
Else
Text1(0).ForeColor = &HFF00&
End If
Else
ksjs = 0
Text1(0).ForeColor = &HFF&
End If
If ksjs = 3 Then
Timer2.Enabled = False
flag1 = 4
cll = Format(Val(Text1(3)) / 1000, "#0.00000")   ''''''''''称量单位g转换成kg
sql1 = "UPDATE pldr SET 实际称量='" & cll & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''关闭盖子
Text4 = 0
Text11 = ""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dzbl(1) = 0
dzbl(2) = 0
dzbl(3) = 0
dzbl(4) = 0
       ReDim WriteData(0 To 14) As String  ''''''写入个数
       Dim DataW As String    '*浮点数的中间处理变量；
       Dim Data10(7) As Single   '*浮点数的中间处理变量；
       Dim Buffer(3) As Byte   '*浮点数的中间处理变量；
 
       For i = 0 To 3
       Data10(i) = Val(dzbl(i + 1))
       CopyMemory Buffer(0), Data10(i), 4
       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
       Next
       a = gk528WriteDevice("D600", 8, WriteData())
 RWorder = 6
 RWcomm = True

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''去皮
qpys = 6                                ''''''''延时准备放料筒盖子原为20秒
Timer6.Enabled = True
'Call VQJC
Call Command4_Click
Text1(0).ForeColor = &HFF&
End If
End Sub

Private Sub VQJC()
On Error Resume Next
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT ISNULL(称量标记,'N'),工序名称,染化助库,染化助名称,配料单位,round(配料用量,6),实际称量,次序号,包装数量,设备编号,设备区位 FROM v_pldr_ft WHERE (称量标记<>'Y' OR 称量标记 IS NULL) AND 料单编号='" & Text3 & "' and 染化助库 not like '%助剂%' ORDER BY 工序名称,次序号"
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

SBBH = 0   ''''''''''''''''''''        设备编号
dzdqpd = 0
dzbl(1) = 0
dzbl(2) = 0
dzbl(3) = 0
dzbl(4) = 0
Timer7.Enabled = True

Else
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''根据称量转换串口
If Adodc3.Recordset.Fields(0) <> "Y" Then

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
qpys = 6                                ''''''''延时准备放料筒盖子原为20秒
Timer6.Enabled = True
End If



If Val(Text1(2)) <= 50 And Val(Text1(2)) >= 0.1 Then
Text1(2) = Format(Text1(2), "#0.00")
flag1 = 0
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在1号称放入容器！"
'''''''''''''''''''''''''''''''''''''''''''''''
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 1
dzbl(1) = Text1(2)
dzbl(2) = 0
dzbl(3) = Adodc3.Recordset.Fields(9)
dzbl(4) = 1

Timer7.Enabled = True
''''''''''''''''''''''''''''''''''''''''''''''
Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If

If Val(Text1(2)) <= 100 And Val(Text1(2)) > 50 Then
Text1(2) = Format(Text1(2), "#0.0")
flag1 = 0
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在1号称放入容器！"
'''''''''''''''''''''''''''''''''''''''''''''''
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 1
dzbl(1) = Text1(2)
dzbl(2) = 0
dzbl(3) = Adodc3.Recordset.Fields(9)
dzbl(4) = 1

Timer7.Enabled = True
''''''''''''''''''''''''''''''''''''''''''''''
Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If


If Val(Text1(2)) > 100 Then
Text1(2) = Format(Text1(2), "#0")
flag1 = 1
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在2号称放入容器！"
SBBH = Adodc3.Recordset.Fields(9)   ''''''''''''''''''''        设备编号
dzdqpd = 2
dzbl(1) = 0
dzbl(2) = Text1(2)
dzbl(3) = Adodc3.Recordset.Fields(9)
dzbl(4) = 1

Timer7.Enabled = True

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
              a = gk528ReadDevice("D704", 6)
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

Private Sub Timer6_Timer()
If qpys <= 0 Then    ''''''去皮延时
Timer6.Enabled = False
Call VQJC
End If
qpys = qpys - 1
If qpys = -1 Then
Label4.Caption = "称重完成！！"
Else
Label4.Caption = "请注意放好料筒盖子！！" + Trim(qpys)
End If
End Sub

Private Sub Timer7_Timer()
If SBBH <> dzdq(3) Then
       ReDim WriteData(0 To 14) As String  ''''''写入个数
       Dim DataW As String    '*浮点数的中间处理变量；
       Dim Data10(7) As Single   '*浮点数的中间处理变量；
       Dim Buffer(3) As Byte   '*浮点数的中间处理变量；
 
       For i = 0 To 3
       Data10(i) = Val(dzbl(i + 1))
       CopyMemory Buffer(0), Data10(i), 4
       DataW = Right("00" + Hex(Buffer(3)), 2) + Right("00" + Hex(Buffer(2)), 2) + Right("00" + Hex(Buffer(1)), 2) + Right("00" + Hex(Buffer(0)), 2)
       WriteData(2 * i + 0) = Val("&H" + Right(DataW, 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(DataW, 1, 4))
       Next
       a = gk528WriteDevice("D600", 8, WriteData())
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






