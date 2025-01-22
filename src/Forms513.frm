VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Forms513 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布扫描"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15915
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   12
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "Text2"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Index           =   11
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "Text2"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Index           =   10
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "Text2"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   8
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "Text2"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   "Text2"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TxtReceive 
      Height          =   375
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   70
      Text            =   "Forms513.frx":0000
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtSend 
      Height          =   375
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   69
      Text            =   "Forms513.frx":0007
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer TmrInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9120
      Top             =   9480
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "条码输入"
      Height          =   2775
      Left            =   480
      TabIndex        =   54
      Top             =   3600
      Width           =   5655
      Begin VB.Label Label13 
         BackColor       =   &H0000C0C0&
         Caption         =   "自动"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   68
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000C0C0&
         Caption         =   "手动"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   67
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "清除"
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
         Left            =   3840
         TabIndex        =   66
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "J"
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
         Index           =   10
         Left            =   4680
         TabIndex        =   65
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "9"
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
         Index           =   9
         Left            =   2880
         TabIndex        =   64
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "8"
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
         Index           =   8
         Left            =   1920
         TabIndex        =   63
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "7"
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
         Index           =   7
         Left            =   960
         TabIndex        =   62
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "6"
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
         Index           =   6
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "5"
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
         Index           =   5
         Left            =   4680
         TabIndex        =   60
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "4"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "3"
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
         Index           =   3
         Left            =   2880
         TabIndex        =   58
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "2"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "1"
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
         Index           =   1
         Left            =   960
         TabIndex        =   56
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   6
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
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
      Index           =   5
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "操作"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   6600
      TabIndex        =   32
      Top             =   1320
      Width           =   8175
      Begin VB.OptionButton Option5 
         Caption         =   "开幅"
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
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "交接产量"
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
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Index           =   9
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "交接"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "车工"
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
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "验布"
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
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "交接产量"
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
         Index           =   6
         Left            =   2760
         TabIndex        =   51
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "."
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
         Index           =   10
         Left            =   6600
         TabIndex        =   50
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9"
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
         Index           =   9
         Left            =   5880
         TabIndex        =   49
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "8"
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
         Index           =   8
         Left            =   5160
         TabIndex        =   48
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7"
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
         Index           =   7
         Left            =   4440
         TabIndex        =   47
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6"
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
         Index           =   6
         Left            =   3720
         TabIndex        =   46
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5"
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
         Index           =   5
         Left            =   7320
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4"
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
         Index           =   4
         Left            =   6600
         TabIndex        =   44
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3"
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
         Index           =   3
         Left            =   5880
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
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
         Index           =   2
         Left            =   5160
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
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
         Index           =   0
         Left            =   3720
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "清除"
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
         Left            =   7320
         TabIndex        =   39
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "产量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6600
      TabIndex        =   19
      Top             =   4320
      Width           =   8175
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
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
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
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
         Index           =   1
         Left            =   840
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3"
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
         Index           =   3
         Left            =   2280
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4"
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
         Index           =   4
         Left            =   3000
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5"
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
         Index           =   5
         Left            =   3720
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6"
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
         Index           =   6
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7"
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
         Index           =   7
         Left            =   5160
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "8"
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
         Index           =   8
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9"
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
         Index           =   9
         Left            =   6360
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "."
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
         Index           =   10
         Left            =   6960
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "清除"
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
         Left            =   7560
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   3
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Left            =   480
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   480
      Width           =   5295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "转数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2160
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "."
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
         Index           =   10
         Left            =   3480
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9"
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
         Index           =   9
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "8"
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
         Index           =   8
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7"
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
         Index           =   7
         Left            =   960
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6"
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
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5"
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
         Index           =   5
         Left            =   4320
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "4"
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
         Index           =   4
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "3"
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
         Index           =   3
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "2"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
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
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "清除"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   1200
         Width           =   735
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forms513.frx":000E
      Height          =   2175
      Left            =   480
      TabIndex        =   76
      Top             =   6600
      Width           =   14295
      _cx             =   25215
      _cy             =   3836
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
      GridLines       =   2
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
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forms513.frx":0023
      Height          =   390
      Left            =   7440
      TabIndex        =   77
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "mc"
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
   Begin MSCommLib.MSComm MSComm 
      Left            =   8520
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   1920
      Top             =   10200
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
      Height          =   330
      Left            =   1800
      Top             =   10200
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
      Left            =   1320
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   1560
      Top             =   10200
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
      Left            =   1800
      Top             =   10200
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
      Left            =   2040
      Top             =   10200
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
      Left            =   2280
      Top             =   10200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   2640
      Top             =   10200
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
      Left            =   3000
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Left            =   2640
      Top             =   10200
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
      Bindings        =   "Forms513.frx":0038
      Height          =   495
      Left            =   480
      TabIndex        =   78
      Top             =   8760
      Visible         =   0   'False
      Width           =   14295
      _cx             =   25215
      _cy             =   873
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
      GridLines       =   2
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "单据"
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
      Index           =   10
      Left            =   480
      TabIndex        =   95
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "开幅"
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
      Index           =   9
      Left            =   12000
      TabIndex        =   94
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "交班"
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
      Index           =   8
      Left            =   9360
      TabIndex        =   93
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "班次"
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
      Index           =   5
      Left            =   6600
      TabIndex        =   92
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "编号"
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
      Index           =   4
      Left            =   2280
      TabIndex        =   91
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "验布"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12960
      TabIndex        =   90
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "验布"
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
      Index           =   3
      Left            =   9360
      TabIndex        =   89
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   4320
      TabIndex        =   88
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "机台"
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
      Index           =   2
      Left            =   480
      TabIndex        =   87
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   86
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "报表转数"
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
      Index           =   1
      Left            =   480
      TabIndex        =   85
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "条码扫描区"
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
      Index           =   18
      Left            =   480
      TabIndex        =   84
      Top             =   120
      Width           =   5295
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
      Left            =   480
      TabIndex        =   83
      Top             =   1800
      Width           =   615
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
      Left            =   3600
      TabIndex        =   82
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "车工"
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
      Left            =   6600
      TabIndex        =   81
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "产量"
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
      Index           =   7
      Left            =   8520
      TabIndex        =   80
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "设定转数"
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
      Index           =   21
      Left            =   2280
      TabIndex        =   79
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Forms513"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim SendCount  As Long     '定义已发送字符对应字节数
    Dim ReceiveCount  As Long  '定义已接收字符对应字节数
    Dim PortSwitch As Boolean    '定义串口是否打开标志
    Public L As String
    Dim DisplayFlag As Boolean   '定义接收窗口是否继续显示标志
Dim ID1 As Long
Dim ID2 As Long
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Form_Load()
On Error Resume Next
For i = 0 To 11
Text2(i).Text = ""
Next
Text1.Text = ""
DataCombo1 = "白班"
       TxtReceive.Text = ""
       TxtSend = ""
       MSComm.CommPort = 1
       MSComm.Settings = "1200,n,8,1"
       MSComm.InBufferSize = 512            ' 设置接收缓冲区为1024字节
       MSComm.OutBufferSize = 512           ' 设置发送缓冲区为4096字节
       MSComm.InBufferCount = 0              ' 清空输入缓冲区
       MSComm.OutBufferCount = 0             ' 清空输出缓冲区
       MSComm.SThreshold = 1                 ' 发送缓冲区空触发发送事件
       MSComm.RThreshold = 1                 ' 每X个字符到接收缓冲区引起触发接收事件
If GetID(ID1, ID2, DevicePath) = 0 Then
jmg = Hex(ID1) + Hex(ID2)
Else
jmg = ""
End If
Option2.value = True
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc1.Refresh
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select distinct mc from bc"
Adodc7.Refresh
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select * from v_zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc8.Refresh
Text1.TabIndex = 0
End Sub

Private Sub Label1_Click()
If Option2.value = True Then
Text2(4).Text = ""
End If
If Option1.value = True Then
Text2(7).Text = ""
End If
If Option3.value = True Then
Text2(10).Text = ""
End If
If Option4.value = True Then
Text2(9).Text = ""
End If
If Option5.value = True Then
Text2(11).Text = ""
End If
End Sub

Private Sub Label10_Click()
Text2(3).Text = ""
End Sub

Private Sub Label11_Click()
Text1.Text = ""
End Sub

Private Sub Label12_Click()
TmrInterval.Enabled = False
End Sub

Private Sub Label13_Click()
TmrInterval.Enabled = True
End Sub

Private Sub Label14_Click()
If Text2(0) = "" Or Text2(1) = "" Then
MsgBox ("请输入织号,匹号")
Exit Sub
End If
If Option3.value = True Then
Forms516.Text2(0) = Text2(0)
Forms516.Text2(1) = Text2(1)
Forms516.Text4 = Text2(10)
Forms516.Show
Else
Forms516.Text2(0) = Text2(0)
Forms516.Text2(1) = Text2(1)
Forms516.Text4 = Text2(4)
Forms516.Show
End If
End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case Index
If Option2.value = True Then
Text2(4).Text = Text2(4).Text + Label3(Index).Caption
End If
If Option1.value = True Then
Text2(7).Text = Text2(7).Text + Label3(Index).Caption
End If
If Option4.value = True Then
Text2(9).Text = Text2(9).Text + Label3(Index).Caption
End If
If Option3.value = True Then
Text2(10).Text = Text2(10).Text + Label3(Index).Caption
End If
If Option5.value = True Then
Text2(11).Text = Text2(11).Text + Label3(Index).Caption
End If
End Select
End Sub

Private Sub Label4_Click()
Text2(5).Text = ""
End Sub

Private Sub Label5_Click(Index As Integer)
Select Case Index
       Case Index
Text2(5).Text = Text2(5).Text + Label5(Index).Caption
End Select
End Sub

Private Sub Label6_Click()
On Error Resume Next

If GetID(ID1, ID2, DevicePath) = 0 Then
jmg = Hex(ID1) + Hex(ID2)
Else
jmg = ""
End If

If jmg = "10E7660911AE6DC8" Then

If Text2(0).Text = "" Or Text2(1).Text = "" Or Text2(2).Text = "" Or Text2(5).Text = "" Or Text2(6).Text = "" Then
MsgBox ("数据输入不完整，请重新输入")
Exit Sub
End If

If Val(Text2(3)) > 30 Then
MsgBox ("超出重量范围")
Exit Sub
End If

'If DataCombo1 = "" Then
'MsgBox ("请输入班次")
'Exit Sub
'End If


'Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select * from zjbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
'Adodc4.Refresh
'If Adodc4.Recordset.EOF Then
'MsgBox ("请输入验布信息，才能登记产量")
'Exit Sub
'End If


'Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select * from zbclbbzsjc where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "'"
'Adodc4.Refresh
'If Not Adodc4.Recordset.EOF Then
'If Adodc4.Recordset.Fields(1) < Adodc4.Recordset.Fields(2) + Val(Text2(5).Text) Then
'MsgBox ("超出转数!,重新输入")
'Text2(5).Text = 0
'Exit Sub
'End If
'End If


If Len(Trim(Text2(4).Text)) <> 3 Then
MsgBox ("操作员错误！")
Exit Sub
End If


'If Val(Text2(5).Text) = 0 Then
'MsgBox ("转数不能为零")
'Exit Sub
'End If

ML = Date - 1

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 序号 from zbclbb where 日期 between cast('" & ML & "' as datetime) and cast('" & Date & "' as datetime) ORDER BY 序号 desc"
Adodc3.Refresh

If Adodc3.Recordset.EOF Then
xh = 1
Else
xh = Val(Adodc3.Recordset.Fields(0)) + 1
End If

If Val(Text2(3).Text) <= 0 Then
MsgBox ("重量不稳")
Exit Sub
End If

'cl = Format(Val(Text2(5)) / Val(Text2(2)) * Val(Text2(3)), "#0.00")
cl = Val(Text2(3))

If cl > 0 Then
If Text2(10) <> "" And Val(Text2(9)) > 0 Then
cl1 = cl - Val(Text2(9))
sql1 = "insert into zbclbb(班次,织号,匹号,转数,产量,操作员,机台,质检,序号,日期,疵布,编号,开幅,单据) VALUES('" & DataCombo1 & "','" & Text2(0) & "','" & Text2(1) & "','" & Text2(5) & "','" & Text2(9) & "','" & Text2(10) & "','" & Text2(6) & "','" & Text2(7) & "','" & xh & "','" & Now & "',0,'" & Text2(8) & "','" & Text2(11) & "','" & Text2(12) & "')"
sql2 = "insert into zbclbb(班次,织号,匹号,转数,产量,操作员,机台,质检,序号,日期,疵布,编号,开幅,单据) VALUES('" & DataCombo1 & "','" & Text2(0) & "','" & Text2(1) & "','" & Text2(5) & "','" & cl1 & "','" & Text2(4) & "','" & Text2(6) & "','" & Text2(7) & "','" & xh & "','" & Now & "',0,'" & Text2(8) & "','" & Text2(11) & "','" & Text2(12) & "')"
sql3 = "update clbbkpd  set 单号=单,客户=客,支数=支,品名=品,寸数=寸,款号=款,幅宽=幅,克重=克 where 织号='" & Text2(0) & "' and 匹号='" & Text2(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "insert into zbclbb(班次,织号,匹号,转数,产量,操作员,机台,质检,序号,日期,疵布,编号,开幅,单据) VALUES('" & DataCombo1 & "','" & Text2(0) & "','" & Text2(1) & "','" & Text2(5) & "','" & cl & "','" & Text2(4) & "','" & Text2(6) & "','" & Text2(7) & "','" & xh & "','" & Now & "',0,'" & Text2(8) & "','" & Text2(11) & "','" & Text2(12) & "')"
sql2 = "update clbbkpd  set 单号=单,客户=客,支数=支,品名=品,寸数=寸,款号=款,幅宽=幅,克重=克 where 织号='" & Text2(0) & "' and 匹号='" & Text2(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If


Adodc6.RecordSource = "select * from mprk  WHERE 织号='" & Text2(0) & "' and 匹号='" & Text2(1) & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
sql4 = "INSERT INTO MPRK(客户,单号,织号,款号,品名,支数,寸数,匹号,重量,幅宽,克重,备注,序号,日期,保管,单价) SELECT 加工,单号,织号,款号,品名,支数,寸数,匹号,'" & Text2(3) & "',幅宽,克重,'','" & xh & "','" & Now & "','" & yhm & "',单价  FROM v_sm_kpd_zz_bc WHERE 织号='" & Text2(0) & "' and 匹号='" & Text2(1) & "'"
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
End If
End If
End If

Adodc1.RecordSource = "select * from zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc1.Refresh
Adodc8.RecordSource = "select * from v_zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc8.Refresh

Text2(9) = ""
Text2(10) = ""
For i = 2 To 6
Text2(i).Text = ""
Next
Text1.SetFocus
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label8_Click(Index As Integer)
Select Case Index
       Case Index
Text1.Text = Text1.Text + Label8(Index).Caption
End Select
End Sub

Private Sub Label9_Click(Index As Integer)
Select Case Index
       Case Index
Text2(3).Text = Text2(3).Text + Label9(Index).Caption
End Select
End Sub

Private Sub Text1_Change()
On Error Resume Next

Dim ZHT  As Long

If InStr(Text1.Text, "J") > 0 Then
ZHT = Val(Mid(Text1.Text, 1, Len(Text1.Text) - 1))

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select 织号,起始,起号,机台,编号,isnull(单据,'') from jhscbq where '" & ZHT & "' between 起号 and 结号"
Adodc5.Refresh

If Not Adodc5.Recordset.EOF Then

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select '' as 车台,转数 from zbkpd where 织号='" & Adodc5.Recordset.Fields(0) & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
For i = 0 To 6
Text2(i).Text = ""
Next
Text2(8).Text = ""
Text2(12) = ""
Text1.Text = ""
Text1.SetFocus
Else
Adodc2.Recordset.MoveFirst
Text2(0).Text = Adodc5.Recordset.Fields(0)
Text2(1).Text = Adodc5.Recordset.Fields(1) + ZHT - Adodc5.Recordset.Fields(2)
Text2(2).Text = Adodc2.Recordset.Fields(1)
Text2(12).Text = Adodc5.Recordset.Fields(5)   ''''''单据
Adodc1.RecordSource = "select * from zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc1.Refresh
Adodc8.RecordSource = "select * from v_zbclbb where 织号='" & Text2(0).Text & "' and 匹号='" & Text2(1).Text & "' order by 序号 desc"
Adodc8.Refresh

If Not Adodc1.Recordset.EOF Then
Text2(5).Text = Adodc2.Recordset.Fields(1) - Adodc1.Recordset.Fields(20)
Else
Text2(5).Text = Adodc2.Recordset.Fields(1)
End If
Text2(6).Text = Adodc5.Recordset.Fields(3)
Text2(8).Text = Adodc5.Recordset.Fields(4)
Text1.Text = ""
End If
Text1.Text = ""
Else
Text1.Text = ""
Text2(0).Text = ""
Text2(1).Text = ""
Text2(2).Text = ""
Text2(5).Text = ""
Text2(6).Text = ""
Text2(8).Text = ""
Text2(12) = ""
End If
End If
End Sub

Private Sub Text2_Change(Index As Integer)
Select Case Index
Case 3
If Val(Text2(3)) > 30 Then
Text2(3) = ""
End If
End Select
End Sub

Private Sub TmrInterval_Timer()
On Error Resume Next
If MSComm.PortOpen = False Then
            MSComm.PortOpen = True
        End If
MSComm.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until MSComm.InBufferCount >= 12
a = MSComm.Input
Text2(3) = Format(Val(Mid(a, 3, 7)), "#0.0")
End Sub


