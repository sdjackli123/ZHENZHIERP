VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Formr330 
   Caption         =   "粉体半自动称量系统"
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
      TabPicture(0)   =   "Formr330.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "配料信息"
      TabPicture(1)   =   "Formr330.frx":001C
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
         TabIndex        =   24
         Top             =   600
         Width           =   18495
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "通讯口操作："
            Height          =   1335
            Index           =   0
            Left            =   10200
            TabIndex        =   109
            Top             =   360
            Width           =   6615
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   2880
               TabIndex        =   113
               Text            =   "Text6"
               Top             =   840
               Width           =   1575
            End
            Begin VB.CommandButton Command11 
               BackColor       =   &H00C0C0FF&
               Caption         =   "关闭串口"
               Height          =   375
               Left            =   5160
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H00C0C0FF&
               Caption         =   "打开串口"
               Height          =   375
               Left            =   5160
               Style           =   1  'Graphical
               TabIndex        =   111
               Top             =   240
               Width           =   1095
            End
            Begin VB.ComboBox Combo1 
               Height          =   300
               ItemData        =   "Formr330.frx":0038
               Left            =   240
               List            =   "Formr330.frx":003A
               TabIndex        =   110
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
               TabIndex        =   116
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
               TabIndex        =   115
               Top             =   840
               Width           =   900
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "端口号："
               Height          =   180
               Left            =   240
               TabIndex        =   114
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "实时读Y0--Y7"
            Height          =   1935
            Index           =   0
            Left            =   10200
            TabIndex        =   62
            Top             =   1800
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
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
               TabIndex        =   80
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
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
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
               TabIndex        =   75
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
               TabIndex        =   74
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
               TabIndex        =   73
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
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   108
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
               TabIndex        =   107
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
               TabIndex        =   106
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   103
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
               TabIndex        =   102
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
               TabIndex        =   101
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
               TabIndex        =   100
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
               TabIndex        =   99
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
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
               TabIndex        =   87
               Top             =   1080
               Width           =   315
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "数值操作"
            Height          =   2775
            Left            =   10200
            TabIndex        =   43
            Top             =   3840
            Visible         =   0   'False
            Width           =   6615
            Begin VB.TextBox Text7 
               Height          =   390
               Left            =   3840
               TabIndex        =   57
               Text            =   "Text7"
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox Text5 
               Height          =   375
               Left            =   960
               TabIndex        =   56
               Text            =   "Text5"
               Top             =   1680
               Width           =   1215
            End
            Begin VB.CommandButton Command9 
               Caption         =   "写入"
               Height          =   420
               Left            =   4440
               TabIndex        =   55
               Top             =   960
               Width           =   975
            End
            Begin VB.CommandButton Command8 
               Caption         =   "读值"
               Height          =   420
               Left            =   3840
               TabIndex        =   54
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
               TabIndex        =   53
               Text            =   "text8"
               Top             =   2160
               Width           =   1575
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   240
               TabIndex        =   49
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton Option13 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   52
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option12 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   51
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "D"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   50
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
               TabIndex        =   45
               Top             =   960
               Width           =   2895
               Begin VB.OptionButton Option14 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "浮点"
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   48
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option5 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "32位"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   47
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "16位"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   46
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   960
               TabIndex        =   44
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
               TabIndex        =   58
               Top             =   2160
               Width           =   720
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "状态操作"
            Height          =   2175
            Left            =   10200
            TabIndex        =   29
            Top             =   6720
            Visible         =   0   'False
            Width           =   6615
            Begin VB.CommandButton Command7 
               Caption         =   "查询当前状态"
               Height          =   420
               Left            =   3120
               TabIndex        =   40
               Top             =   1560
               Width           =   1335
            End
            Begin VB.CommandButton Command6 
               Caption         =   "置位"
               Height          =   420
               Left            =   360
               TabIndex        =   39
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton Command1 
               Caption         =   "复位"
               Height          =   420
               Left            =   1800
               TabIndex        =   38
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
               TabIndex        =   37
               Text            =   "text4"
               Top             =   960
               Width           =   1575
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00C0FFC0&
               Caption         =   "元件选择"
               Height          =   615
               Left            =   345
               TabIndex        =   30
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton Option10 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "S"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   36
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option8 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "C"
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   35
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option9 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   34
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option7 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "M"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   33
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton Option6 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Y"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   32
                  Top             =   240
                  Width           =   495
               End
               Begin VB.OptionButton Option11 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "X"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   31
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
               TabIndex        =   42
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
               TabIndex        =   41
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
            TabIndex        =   26
            Top             =   600
            Width           =   3255
            Begin VB.OptionButton Option2 
               BackColor       =   &H0000FF00&
               Caption         =   "已称量"
               Height          =   495
               Left            =   1680
               TabIndex        =   28
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H000000FF&
               Caption         =   "未称量"
               Height          =   495
               Left            =   240
               TabIndex        =   27
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
            TabIndex        =   25
            Top             =   720
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1680
            TabIndex        =   117
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   255
            Format          =   329449473
            CurrentDate     =   36892
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1680
            TabIndex        =   118
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   329449473
            CurrentDate     =   36892
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Bindings        =   "Formr330.frx":003C
            Height          =   6375
            Left            =   480
            TabIndex        =   119
            Top             =   2520
            Width           =   8295
            _cx             =   14631
            _cy             =   11245
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            FormatString    =   $"Formr330.frx":0051
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
            TabIndex        =   122
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
            TabIndex        =   121
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
            TabIndex        =   120
            Top             =   1200
            Width           =   1335
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
            Text            =   "Text2"
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
            Left            =   9360
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
            Left            =   9360
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
            Left            =   9360
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
            Left            =   9360
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   6600
            Width           =   5655
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
            Bindings        =   "Formr330.frx":0128
            Height          =   5535
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   7695
            _cx             =   13573
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
         Begin VB.Label Label11 
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
            Left            =   8160
            TabIndex        =   123
            Top             =   7320
            Width           =   1215
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
            TabIndex        =   23
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
            Left            =   10920
            TabIndex        =   22
            Top             =   3600
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
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
            TabIndex        =   21
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "条码扫描"
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            Left            =   8160
            TabIndex        =   18
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
            Left            =   8160
            TabIndex        =   17
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
            Left            =   8160
            TabIndex        =   16
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
            Height          =   615
            Index           =   3
            Left            =   8160
            TabIndex        =   15
            Top             =   6600
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   600
            Width           =   4215
         End
      End
   End
End
Attribute VB_Name = "Formr330"
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



Private Sub Command4_Click()
If Option1.value = True Then
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and (称量标记='N' or 称量标记 is null) AND cast(CONVERT(varchar(120),配料日期,23) as datetime)  between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and 称量标记='Y' AND cast(CONVERT(varchar(120),配料日期,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
End If
VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(4) = 2500
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
DTPicker1.value = Date
DTPicker2.value = Date

Label4.Caption = ""

MSComm1.CommPort = 1
MSComm1.Settings = "600,n,8,1"

MSComm2.CommPort = 2
MSComm2.Settings = "600,n,8,1"

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


  Dim b As String
  
  b = OpenComm(MSComm4, 3, "9600,e,7,1")
  
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
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and (称量标记='N' or 称量标记 is null) AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT distinct 锅号,重量,料单编号,配料日期,称量标记 FROM pldr WHERE 染化助库<>'助剂库' and 称量标记='Y' AND 配料日期 between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) ORDER BY 料单编号"
Adodc1.Refresh
End If



Text2.TabIndex = 0
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(3) = 2500

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(4) = 2500

End Sub



Private Sub Label10_Click()
Text3 = ""
Text2 = ""
Text2.SetFocus
End Sub

Private Sub Label11_Click()
If MSComm1.PortOpen = True Then
MSComm1.Output = Chr$(27) + "t"
End If

If MSComm2.PortOpen = True Then
MSComm2.Output = Chr$(27) + "t"
End If
End Sub

Private Sub Label8_Click()
cll = Text1(2)
sql1 = "UPDATE pldr SET 实际称量='" & cll & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''关闭盖子
Text4 = 1

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
    
   b = MSCONComm(MSComm4)
   Text6.Text = b
   If b <> "0" Then Exit Sub
   Timer4.Enabled = False
   Select Case Order
          Case 0   'read d0--d6
   '
                '   Ddata(0) = "&H" + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2)
                            '*PLC返回的寄存器数值是从低字节到高字节排列，所以我们要重新排列一下！
                '   Text1(4).Text = CStr(Val(Ddata(0)))
               
          Case 5   '临时读
               If Option4.value = True Then  '16位
                  Ddata(0) = "&H" + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2) '*PLC返回的寄存器数值是从低字节到高字节排列，所以我们要重新排列一下！
                  Text8.Text = CStr(Val(Ddata(0)))
               Else
                  If Option5.value = True Then '32位
                     Ddata(0) = "&H" + Mid(PLCText, 7, 2) + Mid(PLCText, 5, 2) + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2) '*PLC返回的寄存器数值是从低字节到高字节排列，所以我们要重新排列一下！
                     Text8.Text = CStr(Val(Ddata(0)))
                  Else  '浮点数
                      Dim Data10 As Long    '*浮点数中间处理变量；
                      Dim Data As Single    '*浮点数中间处理变量；
                      Dim dataCl As String  '*浮点数中间处理变量；
                      dataCl = Mid(PLCText, 7, 2) + Mid(PLCText, 5, 2) + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2)
                      Data10 = Val("&H" & dataCl) '*转换为十进制
                      CopyMemory Data, Data10, 4 '*转换为浮点数，调用模块中的COPY声明,意义为:将L复制给F1,位数为四位.
                      Text8.Text = CStr(Data)
                  End If
               End If
               Order = 0
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
flag3 = False
End If

If Text1(4) = "1" Then
Label4.Caption = "请注意是否去皮！！"
Beep 2000, 50
qpys = 5                                ''''''''延时准备变量为5秒
Timer1.Enabled = True
flag3 = True
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
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from v_ftbzsl where 料单编号='" & Text3 & "' and isnull(包装箱数,0)<>isnull(称量箱数,0)"
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
  a = gk528SetDevice(Adree, 0)  '地址  置位为1 复位为0
  RWorder = 7
  RWcomm = True
End If
If Val(Text4) = 0 Then
  Adree = "M66"
  a = gk528SetDevice(Adree, 1)  '地址  置位为1 复位为0
  RWorder = 7
  RWcomm = True
End If
End Sub

Private Sub Text7_Change()
If Val(Text7) > 100 Then
    Adree = "D" & Text5.Text
    ReDim WriteData(0) As String
    WriteData(0) = Val(Text7.Text)
    a = gk528WriteDevice(Adree, 1, WriteData)   '地址  个数  数值组
 RWorder = 6
 RWcomm = True
Text1(4) = "1"
End If
End Sub

Private Sub Text8_Change()
Text1(4) = Text8
End Sub


Private Sub Timer1_Timer()
If qpys = 1 Then    ''''''去皮延时

If flag1 = 0 Then
If flag3 = True Then
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
End If

If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm1.InBufferCount >= 13
a = MSComm1.Input
Text1(3) = Val(Trim(Mid(a, 1, 9)))            ''''''称重量

If Val(Text1(3)) <> 0 Then    ''''''''''''''''''''''''如果不等于0  去皮
If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "t"
End If

End If
Timer5.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = True
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If flag1 = 1 Then
If flag3 = True Then
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
End If


If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
        End If
MSComm2.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm2.InBufferCount >= 13
a = MSComm2.Input

Text1(3) = Val(Format(Val(Mid(a, 1, 9)), "#0"))

If Val(Text1(3)) <> 0 Then    ''''''''''''''''''''''''如果不等于0  去皮
If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
        End If
MSComm2.Output = Chr$(27) + "t"
End If

End If
Timer5.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = True
End If

End If

qpys = qpys - 1
Label4.Caption = "请注意是否去皮！！" + Trim(qpys)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Sub Timer2_Timer()

If Val(Text1(2)) <= 100 Then
If Val(Text1(3)) > 0 And Val(Text1(2)) <= (Val(Text1(3)) + 0.02) And Val(Text1(2)) >= (Val(Text1(3)) - 0.02) And Val(Text1(2)) > 0 Then  ''''''误差在——+0.02g
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

Else

If Val(Text1(3)) > 0 And Val(Text1(2)) <= (Val(Text1(3)) + 1) And Val(Text1(2)) >= (Val(Text1(3)) - 1) And Val(Text1(2)) > 0 Then ''''''误差在——+1g
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
End If




If ksjs = 3 Then
Timer2.Enabled = False
flag1 = 4
cll = Format(Val(Text1(3)) / 1000, "#0.00000")   ''''''''''称量单位g转换成kg
sql1 = "UPDATE pldr SET 实际称量='" & cll & "',称量员='" & yhm & "',称量标记='Y',称量日期='" & Now & "' WHERE 料单编号='" & Text3 & "' and 染化助名称='" & Text1(0) & "' and 次序号='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Adodc2.RecordSource = "SELECT 工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,次序号 FROM pldr WHERE 料单编号='" & Text3 & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc2.Refresh
''''''''''''''''''''''''''''''''关闭盖子
Text4 = 1
Text11 = ""

'''''''''''''''''''''''''''''''''''''''''
If MSComm1.PortOpen = True Then
MSComm1.Output = Chr$(27) + "t"
End If
If MSComm2.PortOpen = True Then
MSComm2.Output = Chr$(27) + "t"
End If

Call VQJC
Call Command4_Click
Text1(0).ForeColor = &HFF&
End If
End Sub

Private Sub VQJC()
On Error Resume Next
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "SELECT ISNULL(称量标记,'N'),工序名称,染化助库,染化助名称,配料单位,round(配料用量,6),实际称量,次序号,包装数量,设备编号,设备区位 FROM v_pldr_ft WHERE (称量标记<>'Y' OR 称量标记 IS NULL) AND 料单编号='" & Text3 & "' and 染化助库<>'助剂库' ORDER BY 工序名称,次序号"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Label4.Caption = "称重完成"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''称量后置位
Text4 = "1"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select * from v_ftbzsl where 料单编号='" & Text3 & "' and isnull(包装箱数,0)<>isnull(称量箱数,0)"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
Formr333.Text3 = Text3
Formr333.Show
End If

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

If Val(Text1(2)) <= 50 And Val(Text1(2)) >= 0.1 Then
Text1(2) = Format(Text1(2), "#0.00")
flag1 = 0
flag3 = False
Timer1.Enabled = False
Label4.Caption = "请在1号称放入容器！"
'''''''''''''''''''''''''''''''''''''''''''''''
Text11 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text5 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text7 = Adodc3.Recordset.Fields(9)  ''''''''''''''''''''        设备编号
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
Text11 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text5 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text7 = Adodc3.Recordset.Fields(9)  ''''''''''''''''''''        设备编号
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
Text11 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text5 = Adodc3.Recordset.Fields(10) ''''''''''''''''''''写入PLC 转盘区位
Text7 = Adodc3.Recordset.Fields(9)  ''''''''''''''''''''        设备编号

Text1(3) = 0
Text1(4) = ""
Text1(4).SetFocus
End If

Text4 = 0
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
         Case 0   '读D56
              a = gk528ReadDevice("D56", 1)
         Case 1
              a = gk528ReadDevice("M71", 1)
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
If flag1 = 0 Then

If flag3 = True Then
If MSComm2.PortOpen = True Then
            MSComm2.PortOpen = False
End If

If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
MSComm1.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm1.InBufferCount >= 13
a = MSComm1.Input
'If Mid(A, 10, 1) = "g" Then
Text1(3) = Val(Trim(Mid(a, 1, 9)))     ''''''称重量
'End If

End If

End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If flag1 = 1 Then
If flag3 = True Then
If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
End If


If MSComm2.PortOpen = False Then
            MSComm2.PortOpen = True
        End If
MSComm2.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until flag2 And MSComm2.InBufferCount >= 13
a = MSComm2.Input
'If Mid(A, 10, 2) = "kg" Then
Text1(3) = Val(Format(Val(Mid(a, 1, 9) * 1000), "#0"))
'End If

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




