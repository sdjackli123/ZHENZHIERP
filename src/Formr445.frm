VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formr445 
   BackColor       =   &H00C0E0FF&
   Caption         =   "定型助剂申请"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15945
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11055
      Left            =   -1080
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   19500
      _Version        =   393216
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
      TabCaption(0)   =   "配料信息"
      TabPicture(0)   =   "Formr445.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "称量信息"
      TabPicture(1)   =   "Formr445.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "料单信息"
      TabPicture(2)   =   "Formr445.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2(0)"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0E0FF&
         Height          =   10335
         Left            =   -74520
         ScaleHeight     =   10275
         ScaleWidth      =   18435
         TabIndex        =   41
         Top             =   600
         Width           =   18495
         Begin VB.TextBox Text13 
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
            Index           =   0
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "Text13"
            Top             =   3360
            Width           =   5415
         End
         Begin VB.TextBox Text13 
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
            Index           =   1
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   63
            Text            =   "Text13"
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox Text13 
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
            Index           =   2
            Left            =   9720
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "Text13"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C0E0FF&
            Caption         =   "操作输入键"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4455
            Left            =   7200
            TabIndex        =   49
            Top             =   3960
            Width           =   6255
            Begin VB.Label Label19 
               BackColor       =   &H00FFFFC0&
               Caption         =   "清除"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   26.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   1440
               TabIndex        =   61
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   10
               Left            =   120
               TabIndex        =   60
               Top             =   3000
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   9
               Left            =   5160
               TabIndex        =   59
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   8
               Left            =   3960
               TabIndex        =   58
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   7
               Left            =   2760
               TabIndex        =   57
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   6
               Left            =   1440
               TabIndex        =   56
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   5
               Left            =   120
               TabIndex        =   55
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   4
               Left            =   5160
               TabIndex        =   54
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   3
               Left            =   3960
               TabIndex        =   53
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   2
               Left            =   2760
               TabIndex        =   52
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   1
               Left            =   1440
               TabIndex        =   51
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFC0&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   42
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Index           =   14
               Left            =   120
               TabIndex        =   50
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.OptionButton Option18 
            Caption         =   "料缸编号输入"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   7200
            TabIndex        =   48
            Top             =   3120
            Width           =   2295
         End
         Begin VB.OptionButton Option18 
            Caption         =   "定型布量输入"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   10320
            TabIndex        =   47
            Top             =   3120
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "是否包含水槽水量"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   11280
            TabIndex        =   44
            Top             =   840
            Width           =   2535
            Begin VB.OptionButton Option16 
               BackColor       =   &H00FFFFC0&
               Caption         =   "是"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   24
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   240
               TabIndex        =   46
               Top             =   480
               Width           =   855
            End
            Begin VB.OptionButton Option17 
               BackColor       =   &H0000C0C0&
               Caption         =   "否"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   24
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   1440
               TabIndex        =   45
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.TextBox Text13 
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
            Index           =   3
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "Text13"
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox Text13 
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
            Index           =   4
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Text13"
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送料缸编号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6840
            TabIndex        =   84
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送工艺信息"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   720
            TabIndex        =   83
            Top             =   2640
            Width           =   5415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   0
            Left            =   720
            TabIndex        =   82
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Left            =   720
            TabIndex        =   81
            Top             =   4920
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   2
            Left            =   720
            TabIndex        =   80
            Top             =   5760
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Left            =   720
            TabIndex        =   79
            Top             =   6600
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   4
            Left            =   720
            TabIndex        =   78
            Top             =   7440
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   5
            Left            =   720
            TabIndex        =   77
            Top             =   8280
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   6
            Left            =   720
            TabIndex        =   76
            Top             =   9120
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   7
            Left            =   3840
            TabIndex        =   75
            Top             =   4080
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   8
            Left            =   3840
            TabIndex        =   74
            Top             =   4920
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   9
            Left            =   3840
            TabIndex        =   73
            Top             =   5760
            Width           =   2415
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0FFC0&
            Caption         =   "输送确认"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   14040
            TabIndex        =   72
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0FFC0&
            Caption         =   "布量"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9720
            TabIndex        =   71
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   10
            Left            =   3840
            TabIndex        =   70
            Top             =   6600
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Left            =   3840
            TabIndex        =   69
            Top             =   7440
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   12
            Left            =   3840
            TabIndex        =   68
            Top             =   9120
            Width           =   2415
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
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
            Index           =   13
            Left            =   3840
            TabIndex        =   67
            Top             =   8280
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "锅号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   7
            Left            =   720
            TabIndex        =   66
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "品名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   3360
            TabIndex        =   65
            Top             =   840
            Width           =   3255
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   10215
         Index           =   0
         Left            =   -74900
         ScaleHeight     =   10155
         ScaleWidth      =   18435
         TabIndex        =   31
         Top             =   600
         Width           =   18495
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   2
            Left            =   8520
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "Text11"
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   10440
            TabIndex        =   39
            Text            =   "Text14"
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   1
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "Text11"
            Top             =   120
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Index           =   0
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "Text11"
            Top             =   0
            Visible         =   0   'False
            Width           =   615
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
            Left            =   10920
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   3960
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
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
            Left            =   12960
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
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
            Left            =   10920
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2880
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
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
            Left            =   11520
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   5760
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
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
            Left            =   10920
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   4680
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0E0FF&
         Height          =   10455
         Index           =   1
         Left            =   1080
         ScaleHeight     =   10395
         ScaleWidth      =   15675
         TabIndex        =   1
         Top             =   600
         Width           =   15735
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   26.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2880
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   480
            Width           =   3135
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "退出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   14160
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "查询条件"
            Height          =   1455
            Left            =   10080
            TabIndex        =   16
            Top             =   120
            Width           =   2415
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "料缸"
               Height          =   495
               Index           =   9
               Left            =   240
               TabIndex        =   20
               Top             =   840
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "日期"
               Height          =   495
               Index           =   6
               Left            =   240
               TabIndex        =   19
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "待配送"
               Height          =   495
               Index           =   0
               Left            =   1320
               TabIndex        =   18
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00FFFFC0&
               Caption         =   "已配送"
               Height          =   495
               Index           =   1
               Left            =   1320
               TabIndex        =   17
               Top             =   840
               Width           =   855
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "查询"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   12600
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "料缸输入"
            Height          =   1095
            Left            =   240
            TabIndex        =   3
            Top             =   1560
            Width           =   15015
            Begin VB.Label Label1 
               Caption         =   "0"
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
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "1"
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
               Index           =   1
               Left            =   1560
               TabIndex        =   13
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "2"
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
               Index           =   2
               Left            =   2880
               TabIndex        =   12
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "3"
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
               Index           =   3
               Left            =   4200
               TabIndex        =   11
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "4"
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
               Index           =   4
               Left            =   5520
               TabIndex        =   10
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "5"
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
               Index           =   5
               Left            =   6840
               TabIndex        =   9
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "6"
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
               Index           =   6
               Left            =   8160
               TabIndex        =   8
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "7"
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
               Index           =   7
               Left            =   9480
               TabIndex        =   7
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "8"
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
               Index           =   8
               Left            =   10800
               TabIndex        =   6
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "9"
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
               Index           =   9
               Left            =   12120
               TabIndex        =   5
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label4 
               Caption         =   "清除"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   13320
               TabIndex        =   4
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8880
            TabIndex        =   2
            Text            =   "Text3"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Timer Timer1 
            Interval        =   3000
            Left            =   12480
            Top             =   0
         End
         Begin MSAdodcLib.Adodc Adodc9 
            Height          =   330
            Left            =   2520
            Top             =   9720
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
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
         Begin MSAdodcLib.Adodc Adodc8 
            Height          =   330
            Left            =   2640
            Top             =   9600
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
            Height          =   330
            Left            =   2880
            Top             =   9600
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
            Left            =   3600
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
            Left            =   3840
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
            Left            =   8520
            Top             =   9480
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
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Bindings        =   "Formr445.frx":0054
            Height          =   5535
            Left            =   240
            TabIndex        =   23
            Top             =   2880
            Width           =   15975
            _cx             =   28178
            _cy             =   9763
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
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   7080
            TabIndex        =   24
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
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
            Left            =   7080
            TabIndex        =   25
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarTitleBackColor=   8421376
            CalendarTrailingForeColor=   1118719
            Format          =   329318401
            CurrentDate     =   36892
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
            Height          =   735
            Index           =   2
            Left            =   1680
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFF00&
            Caption         =   "重新扫描"
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
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000C0C0&
            Caption         =   "起始"
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
            Left            =   6360
            TabIndex        =   28
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000C0C0&
            Caption         =   "结束"
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
            Left            =   6360
            TabIndex        =   27
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H0000C0C0&
            Caption         =   "料缸"
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
            Index           =   0
            Left            =   8880
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "Formr445"
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
Dim PDCLWB As Integer ''''判断称料完毕
Dim qpys  As Integer    '''''去皮延时
'''''''''''''''''             PLC 变量
Dim GS As Integer   '''锅数
Dim YMSCT As String '位元件操作选择标志
Dim Adree As String ' 元件地址
Dim Order As Integer '通讯顺序
Dim RWorder As Integer ' 读写通讯顺序
Dim RWcomm As Boolean '读取操作
Dim ysbc As Integer '''''''寄存器延时保持
Dim SJPD As Integer
Dim dqdz As Integer ''''''''判断是否数据
Dim ytmd As Double ''''液体密度
Dim dczw1, dczw2, dczw3, dczw4, dczw5, dczw6 As Integer ''''''''判断是否有称量数据
Dim bcbl1, bcbl2, bcbl3 As Integer ''''''''数据保存
Dim xrld, xrld1, xrld2, xrld3 As Integer ''''''''写入料单信息
Dim ytsz(7) As String ''''''''''''''''''''液体变量数组
Dim ztdq1(4) As String ''''''''''''''''''''1号称称量状态读取数组
Dim ztdq2(4) As String ''''''''''''''''''''2号称称量状态读取数组
Dim ztdq3(4) As String ''''''''''''''''''''称量数据保存成功数组
Dim ztdq4(4) As String ''''''''''''''''''''4号称称量状态读取数组
Dim ztdq5(4) As String ''''''''''''''''''''m103--m106的状态值
Dim ctbh As String    ''''''''''''''''''''称台编号
Dim czbc As String   '''''''''''''''''''''称量保持数据
Dim wdbl As String   '''''''''''''''''''''称量读取的稳定变量
Dim ssxscsData(35) As Single   ''''''实时显示变量1
Dim csfh  As Integer ''''''''''''''''''''传输液位库存
Dim ssxsData(35) As Single
Dim csfhdz(35)  As Integer   ''''''''''''''''''''传输液位寄存器
Dim plcxryc As Integer     ''''''''延时写入
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   '浮点数处理
Dim MXH  As Integer    '''''''''循环读M


Private Sub Command1_Click()
sql1 = ""

If Check2(0).value = 1 Then
sql1 = sql1 + "isnull(输送时长,0)=0 and "
End If

If Check2(1).value = 1 Then
sql1 = sql1 + "isnull(输送时长,0)>0 and "
End If

If Check2(9).value = 1 Then
sql1 = sql1 + "机台 like '%'+'" & Text3 & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "cast(CONVERT(varchar(120),申请时间,23) as datetime) between cast('" & DTPicker1.value & "' as datetime) and cast('" & DTPicker2.value & "' as datetime) and "
End If

If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)

Adodc1.RecordSource = "SELECT 料单编号,工序名称,染化助库,染化助名称,配料单位,round(配料用量,4) as 配料用量,实际称量,次序号,机台,锅号,申请时间,开始输送,管道编号,车台编号,输送状态,输送时长 FROM v_pldr_dx WHERE 申请时间 is not null and 是否审核='是'  and (" + sql1 + ") AND isnull(管道编号,'')<>'' ORDER BY 申请时间 desc,开始输送 desc,工序名称,次序号"
Adodc1.Refresh

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 500
Next
End If

VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 1000
VSFlexGrid1.ColWidth(2) = 900
VSFlexGrid1.ColWidth(3) = 900
VSFlexGrid1.ColWidth(4) = 1000
VSFlexGrid1.ColWidth(5) = 600
VSFlexGrid1.ColWidth(6) = 1000
VSFlexGrid1.ColWidth(7) = 1000
VSFlexGrid1.ColWidth(8) = 600
VSFlexGrid1.ColWidth(9) = 600
VSFlexGrid1.ColWidth(10) = 1000
VSFlexGrid1.ColWidth(11) = 1800
VSFlexGrid1.ColWidth(12) = 1800
VSFlexGrid1.ColWidth(13) = 600
VSFlexGrid1.ColWidth(14) = 600
VSFlexGrid1.ColWidth(15) = 600
VSFlexGrid1.ColWidth(16) = 600

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
On Error Resume Next
DTPicker1.value = Date - 1
DTPicker2.value = Date


Text17 = ""
Text3 = ""
Option1.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset


Check2(6).value = 1
Check2(9).value = 1
Text3 = 0
For i = 0 To 4
Text13(i) = ""
Next
plcxryc = 1
For i = 0 To 13
Label17(i).Visible = False
Next

Option17.value = True
Option18(0).value = True
For m = 0 To 6
Text1(m) = ""
Next
csfh = 1     '''''''''''''传输发送  液位库存
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
wdbl = "0"


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


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

Private Sub Label1_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
Text3.Text = Label1(Index).Caption
End Select
End Sub

Private Sub Label10_Click()
Text2 = ""
Text2.SetFocus
End Sub

Private Sub Label15_Click(Index As Integer)
On Error Resume Next
Select Case Index
       Case Index
If Option18(0).value = True Then
Text13(1).Text = Label15(Index).Caption
End If

If Option18(1).value = True Then
Text13(2).Text = Text13(2).Text + Label15(Index).Caption
End If
End Select
End Sub

Private Sub Label16_Click()
On Error Resume Next
Adodc5.RecordSource = "select distinct 模板编号 from SSgy order by 模板编号"
Adodc5.Refresh

For i = 0 To 13
Label17(i).Visible = False
Next
If Not Adodc5.Recordset.EOF Then
Adodc5.Recordset.MoveFirst
L = 0
Do While Not Adodc5.Recordset.EOF
Label17(L).Caption = Adodc5.Recordset.Fields(0)
Label17(L).Visible = True
Adodc5.Recordset.MoveNext
L = L + 1
Loop
End If
End Sub

Private Sub Label17_Click(Index As Integer)
Select Case Index
       Case Index
       Text13(0) = Label17(Index).Caption
End Select
End Sub

Private Sub Label18_Click()
If Val(Text13(2)) <= 0 Then
Text13(2) = ""
Exit Sub
End If

If Trim(Text13(1)) = "" Then
Exit Sub
End If

If MsgBox("按照工艺 " + Text13(0) + " 生成输送料单吗？", vbYesNo) = vbNo Then Exit Sub
If Text13(0) = "" Then
MsgBox ("请选择工艺!")
Exit Sub
End If

Adodc6.RecordSource = "SELECT  ISNULL(MAX(RIGHT(料单编号,9)),CONVERT(varchar(100), GETDATE(), 12)+'000')+1 from PLDR where RIGHT(料单编号,9) like CONVERT(varchar(100), GETDATE(), 12) + '%' and left(料单编号,1)='D'"
Adodc6.Refresh
ldbh = "D" + Trim(Adodc6.Recordset.Fields(0))
pf = Val(Text13(2))
If Option16.value = True Then
sql1 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',配方*(cast('" & pf & "' as real)*0.3+100)/100,0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%湿布%' and 染化助名称<>'水'"
sql2 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',配方*(cast('" & pf & "' as real)*1+100)/100,0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%干布%' and 染化助名称<>'水'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc8.RecordSource = "select isnull(sum(配料用量),0) from pldr where 料单编号='" & ldbh & "'"
Adodc8.Refresh
plyl = Val(Adodc8.Recordset.Fields(0))
sql1 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',cast('" & pf & "' as real)*0.3+100-cast('" & plyl & "' as real),0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%湿布%' and 染化助名称='水'"
sql2 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',cast('" & pf & "' as real)*1+100-cast('" & plyl & "' as real),0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%干布%' and 染化助名称='水'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If


If Option17.value = True Then
sql1 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',配方*(cast('" & pf & "' as real)*0.3)/100,0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%湿布%' and 染化助名称<>'水'"
sql2 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',配方*(cast('" & pf & "' as real)*1)/100,0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%干布%' and 染化助名称<>'水'"

RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc8.RecordSource = "select isnull(sum(配料用量),0) from pldr where 料单编号='" & ldbh & "'"
Adodc8.Refresh
plyl = Val(Adodc8.Recordset.Fields(0))
sql1 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',cast('" & pf & "' as real)*0.3-cast('" & plyl & "' as real),0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%湿布%' and 染化助名称='水'"
sql2 = "insert into pldr(锅号,重量,品名,料单编号,工序名称,染化助库,染化助名称,配料单位,配料用量,实际称量,配料日期,次序号,机台,申请时间,是否审核) select '" & Text13(3) & "','" & Text13(2) & "','" & Text13(4) & "','" & ldbh & "','定型工序',染化助库,染化助名称,'kg',cast('" & pf & "' as real)*1-cast('" & plyl & "' as real),0,'" & Date & "',序号,'" & Text13(1) & "','" & Now & "','是' from ssgy  WHERE 模板编号='" & Text13(0) & "' and 模板编号 like '%干布%' and 染化助名称='水'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If

For i = 0 To 4
Text13(i) = ""
Next

If Val(ztdq1(1)) = 0 Then
Text9 = "定型工序"
SSTab1.Tab = 0
Else
SSTab1.Tab = 0
End If
End Sub

Private Sub Label19_Click()
On Error Resume Next
If Option18(0).value = True Then
Text13(1).Text = Mid(Text13(1), 1, Len(Text13(1)) - 1)
End If

If Option18(1).value = True Then
Text13(2).Text = Mid(Text13(2), 1, Len(Text13(2)) - 1)
End If

End Sub






Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 2
pmbl = 5
Formy80.Show
End Select
End Sub

Private Sub Label4_Click()
Text3 = ""
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_Change()
       GS = 0
If InStr(Text2, "J") > 0 Or InStr(Text2, "j") > 0 Then
gh = Mid(Text2, 1, Len(Text2) - 1)
Adodc7.RecordSource = "select 工艺,重量,品名,锅号 from DXGYXX where 锅号='" & gh & "'"
Adodc7.Refresh
If Adodc7.Recordset.EOF Then
MsgBox ("定型工艺没有设定！")
Text13(0) = ""
Text13(2) = ""
Text13(3) = ""
Text13(4) = ""
Text2 = ""
Text2.SetFocus
Else
Text13(0) = Adodc7.Recordset.Fields(0)
Text13(3) = Adodc7.Recordset.Fields(3)
Text13(4) = Adodc7.Recordset.Fields(2)
Adodc3.RecordSource = "select distinct 重量,料单编号 from pldr where 锅号='" & gh & "' and 料单编号 like 'D%'"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
czl = 0
Else
Adodc3.Recordset.MoveFirst
Do While Not Adodc3.Recordset.EOF
czl = czl + Val(Adodc3.Recordset.Fields(0))
Adodc3.Recordset.MoveNext
Loop
End If

If InStr(Text13(0), "干布") > 0 Then
If (Val(Adodc7.Recordset.Fields(1)) - Val(czl)) >= 300 Then
If (Val(Adodc7.Recordset.Fields(1)) / 300) = Int(Val(Adodc7.Recordset.Fields(1)) / 300) Then
GS = Int(Val(Adodc7.Recordset.Fields(1)) / 300)
Text13(2) = Format(Val(Adodc7.Recordset.Fields(1)) / GS, "#0.0")
Else
GS = Int(Val(Adodc7.Recordset.Fields(1)) / 300) + 1
Text13(2) = Format(Val(Adodc7.Recordset.Fields(1)) / GS, "#0.0")
End If
Else
Text13(2) = Format((Val(Adodc7.Recordset.Fields(1)) - Val(czl)), "#0.0")
End If
Text2 = ""
SSTab1.Tab = 1
End If

If InStr(Text13(0), "湿布") > 0 Then
If (Val(Adodc7.Recordset.Fields(1)) - Val(czl)) >= 900 Then
If (Val(Adodc7.Recordset.Fields(1)) / 900) = Int(Val(Adodc7.Recordset.Fields(1)) / 900) Then
GS = Int(Val(Adodc7.Recordset.Fields(1)) / 900)
Text13(2) = Format(Val(Adodc7.Recordset.Fields(1)) / GS, "#0.0")
Else
GS = Int(Val(Adodc7.Recordset.Fields(1)) / 900) + 1
Text13(2) = Format(Val(Adodc7.Recordset.Fields(1)) / GS, "#0.0")
End If
Else
Text13(2) = Format((Val(Adodc7.Recordset.Fields(1)) - Val(czl)), "#0.0")
End If
Text2 = ""
SSTab1.Tab = 1
End If

End If
End If

End Sub

Private Sub Timer1_Timer()
DTPicker1.value = Date - 1
DTPicker2.value = Date
Call Command1_Click
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
r = VSFlexGrid1.RowSel
c = VSFlexGrid1.ColSel
If MsgBox("确定取消配送 料单 " + VSFlexGrid1.TextMatrix(r, 1) + " 吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "delete from pldr WHERE 料单编号='" & VSFlexGrid1.TextMatrix(r, 1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End Sub

