VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form formr450 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����Զ�����ϵͳ"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   15960
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer15 
      Interval        =   3000
      Left            =   11160
      Top             =   0
   End
   Begin VB.Timer Timer14 
      Interval        =   1000
      Left            =   12600
      Top             =   0
   End
   Begin VB.TextBox Text11 
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
      Left            =   14280
      TabIndex        =   80
      Text            =   "Text11"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text10 
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
      Left            =   10680
      TabIndex        =   79
      Text            =   "Text10"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text9 
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
      Left            =   12600
      TabIndex        =   78
      Text            =   "Text9"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   77
      Text            =   "Text7"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   76
      Text            =   "Text7"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Timer Timer13 
      Interval        =   1000
      Left            =   10680
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12120
      Top             =   0
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   75
      Text            =   "Text4"
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   74
      Text            =   "Text4"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10200
      Top             =   0
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Text            =   "Text5"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   72
      Text            =   "Text4"
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   71
      Text            =   "Text4"
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   70
      Text            =   "Text4"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   69
      Text            =   "Text4"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   68
      Text            =   "Text4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   67
      Text            =   "Text4"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer11 
      Interval        =   3000
      Left            =   11640
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   1000
      Left            =   6840
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9720
      Top             =   0
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   66
      Text            =   "Text3"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
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
      Index           =   7
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "Text1"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   7440
      TabIndex        =   64
      Text            =   "Combo1"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   63
      Text            =   "formr450.frx":0000
      Top             =   9360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Index           =   6
      Left            =   14280
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "Text1"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Index           =   4
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "Text1"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "Text1"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   8760
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   8280
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   7320
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9240
      Top             =   0
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ͨѶ�ڲ�����"
      Height          =   1095
      Index           =   0
      Left            =   7560
      TabIndex        =   47
      Top             =   480
      Width           =   4815
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3120
         TabIndex        =   51
         Text            =   "Text6"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�رմ���"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0FF&
         Caption         =   "�򿪴���"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "formr450.frx":0006
         Left            =   240
         List            =   "formr450.frx":0008
         TabIndex        =   48
         Text            =   "COM1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��򿪴���"
         ForeColor       =   &H00000040&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label244 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͨѶ״̬��"
         ForeColor       =   &H00000040&
         Height          =   300
         Index           =   1
         Left            =   3120
         TabIndex        =   53
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˿ںţ�"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ʵʱ��Y0--Y7"
      Height          =   1815
      Index           =   0
      Left            =   11040
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   720
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   720
         Top             =   720
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   480
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   1080
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
         TabIndex        =   40
         Top             =   1080
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   240
         Width           =   315
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
         TabIndex        =   35
         Top             =   240
         Width           =   315
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   240
         Width           =   210
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
         TabIndex        =   25
         Top             =   240
         Width           =   210
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "formr450.frx":000A
      Height          =   3855
      Left            =   840
      TabIndex        =   81
      Top             =   5040
      Width           =   15375
      _cx             =   27120
      _cy             =   6800
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   1320
      Top             =   9480
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   1440
      Top             =   9480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Left            =   1560
      Top             =   9480
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
      Height          =   330
      Left            =   1800
      Top             =   9480
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
      Height          =   330
      Left            =   2040
      Top             =   9480
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
      Left            =   2280
      Top             =   9480
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
   Begin VB.Label Label10 
      Caption         =   "Label10"
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
      Index           =   2
      Left            =   4800
      TabIndex        =   112
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
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
      Left            =   3480
      TabIndex        =   111
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
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
      Index           =   0
      Left            =   2160
      TabIndex        =   110
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ļ���С"
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
      Index           =   10
      Left            =   14280
      TabIndex        =   109
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "˳ʱ����"
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
      Index           =   9
      Left            =   10680
      TabIndex        =   108
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "ƽ��ʱ��"
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
      Index           =   8
      Left            =   9720
      TabIndex        =   107
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "����ʱ��"
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
      Left            =   6240
      TabIndex        =   106
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ϴˮ��"
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
      Left            =   12600
      TabIndex        =   105
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "�س�"
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
      Index           =   5
      Left            =   13920
      TabIndex        =   104
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "��̨"
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
      Left            =   9120
      TabIndex        =   103
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "ֹͣ"
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
      Left            =   8520
      TabIndex        =   102
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "����"
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
      Left            =   7440
      TabIndex        =   101
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "ֹͣ"
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
      Left            =   4080
      TabIndex        =   100
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "����"
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
      Left            =   3000
      TabIndex        =   99
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "������ϴ"
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
      Index           =   4
      Left            =   5280
      TabIndex        =   98
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "������ϴ"
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
      Left            =   840
      TabIndex        =   97
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "����"
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
      Left            =   13920
      TabIndex        =   96
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "ֹͣ"
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
      Left            =   12600
      TabIndex        =   95
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "��ͣ"
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
      Left            =   12600
      TabIndex        =   94
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label5"
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
      Left            =   6000
      TabIndex        =   93
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
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
      Left            =   840
      TabIndex        =   92
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
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
      Index           =   1
      Left            =   840
      TabIndex        =   91
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ͻ�̨���"
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
      Left            =   14280
      TabIndex        =   90
      Top             =   3120
      Width           =   1935
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
      Height          =   495
      Index           =   0
      Left            =   12600
      TabIndex        =   89
      Top             =   3120
      Width           =   1455
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
      Left            =   9240
      TabIndex        =   88
      Top             =   3120
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
      Height          =   495
      Index           =   2
      Left            =   10680
      TabIndex        =   87
      Top             =   3120
      Width           =   1695
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
      Left            =   2880
      TabIndex        =   86
      Top             =   3120
      Width           =   1815
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
      Left            =   4920
      TabIndex        =   85
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʾ��Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   840
      TabIndex        =   84
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   83
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������Ϣ"
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
      Left            =   840
      TabIndex        =   82
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "formr450"
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
Dim yxztbl(10) As Integer '''''''����״̬����
Dim ksjs As Integer      '''''�����ȶ�����
Dim PDCLWB As Integer ''''�жϳ������
Dim qpys  As Integer    '''''ȥƤ��ʱ
Dim txpd As Integer   ''''ͨѶ״̬�ж�
'''''''''''''''''             PLC ����
Dim YMSCT As String 'λԪ������ѡ���־
Dim Adree As String ' Ԫ����ַ
Dim Order As Integer 'ͨѶ˳��
Dim RWorder As Integer ' ��дͨѶ˳��
Dim RWcomm As Boolean '��ȡ����
Dim ysbc As Integer '''''''�ж�ԭ��ʱ��
Dim SJPD As Integer  '''''''�������쳣
Dim dqdz As Integer ''''''''�ļ���С
Dim pdsfjl As Integer ''''�ж��Ƿ����
Dim wdbj As String   ''''�ȶ����
Dim dczw1, dczw2, dczw3, dczw4, dczw5, dczw6 As Integer ''''''''�ж��Ƿ��г�������
Dim bcbl1, bcbl2, bcbl3 As Integer ''''''''���ݱ���
Dim xrld, xrld1, xrld2, xrld3 As Integer ''''''''д���ϵ���Ϣ
Dim ytsz(17) As String ''''''''''''''''''''Һ���������
Dim ztdq1(4) As String ''''''''''''''''''''1�ųƳ���״̬��ȡ����
Dim ztdq2(4) As String ''''''''''''''''''''2�ųƳ���״̬��ȡ����
Dim ztdq3(4) As String ''''''''''''''''''''�������ݱ���ɹ�����
Dim ztdq4(4) As String ''''''''''''''''''''�ϵ������Ϣ
Dim ztdq5(4) As String ''''''''''''''''''''m103--m106��״ֵ̬
Dim xhpdtx, xhpdql, xhpdql1, xhpdql2, xhpdql3, xhpdql4, xhpdql5, xhpdql6, xhpdql7, xhpdql8, cxsfmjc, plcxryc As Integer ''''''''''''''''''''ѭ���ж�ͨѶ״̬
Dim czbc As String   '''''''''''''''''''''������������
Dim wdbl As String   '''''''''''''''''''''������ȡ���ȶ�����
Dim ssxscsData(35) As Single   ''''''ʵʱ��ʾ����1
Dim csfh  As Integer ''''''''''''''''''''����Һλ���
Dim ssxsData(35) As Single
Dim csfhdz(35)  As Integer   ''''''''''''''''''''����Һλ�Ĵ���
Dim sssjxrjc As Integer   ''''''д��PLC
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
   '����������
Dim pdcqplc As Integer
Dim MXH  As Integer    '''''''''ѭ����M
Dim ssgs As Integer   '''''��������
Dim sssl1 As Integer '''''���ʹ���
Dim sssl2 As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
On Error Resume Next
  Dim b As String
  Dim COM1 As Integer
  
  COM1 = Combo1.ListIndex + 1
  b = OpenComm(MSComm4, COM1, "9600,e,7,1")
  
      If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
  
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
 
      If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
 
 Timer3.Enabled = False
 Timer4.Enabled = False
End Sub



Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then
Unload Me
End
End If
sssl1 = 0  '''''�������ʹ���
sssl2 = 0  '''''�������ʹ���
Text11 = ""
plcxryc = 1
For i = 0 To 6
Text1(i) = ""
Next
Text6 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
pdcqplc = 1
Text3 = ""
Label4.Caption = ""
 txpd = 0   '''''ͨѶ�ж�
 pdsfjl = 0
 Text7 = 0
 xhpdql1 = 1
 xhpdql2 = 1
 xhpdql3 = 1
 xhpdql4 = 1
 xhpdql5 = 1
 xhpdql6 = 1
 xhpdql7 = 1
 xhpdql8 = 1
 dqdz = 1
 
Dim g As Integer
      '*���ͨѶ��ѡ�����
      
    For g = 1 To 10                             '*���ͨѶ��ѡ��
        Combo1.AddItem "Com" & CStr(g)
    Next g
    Combo1.ListIndex = 0  '��ʾ��һ��
    
    YMSCT = "M"
    DCT = "D"

  Dim b As String
  
  b = OpenComm(MSComm4, 1, "9600,e,7,1")
  
  If b = 0 Then
     Order = 0
     Timer3.Enabled = True
     Timer4.Enabled = True
     RWcomm = False
 Else
     Timer4.Enabled = False
     Timer3.Enabled = False
 End If


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT �ϵ����,��������,Ⱦ������,Ⱦ��������,���ϵ�λ,round(��������,4) as ��������,ʵ�ʳ���,�����,��̨,�ܵ�����1,�������,�ܵ�����2,�ܵ����,��̨��� FROM v_pldr_dx WHERE ����ʱ�� is not null and �Ƿ����='��' and isnull(�������,'')<>'Y' and isnull(������Ϣ,'') not like '%�쳣����%' and �ܵ����<>0 and (��������-isnull(ʵ�ʳ���,0))>0 ORDER BY ����ʱ��,��������,�����"
Adodc2.Refresh
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(2) = 2000
VSFlexGrid2.ColWidth(3) = 1000
VSFlexGrid2.ColWidth(4) = 2000
VSFlexGrid2.ColWidth(5) = 1000
VSFlexGrid2.ColWidth(6) = 1000
VSFlexGrid2.ColWidth(7) = 1000
VSFlexGrid2.ColWidth(10) = 0
VSFlexGrid2.ColWidth(12) = 0

VSFlexGrid2.RowHeightMin = 600
Me.Hide
End Sub

Private Sub Label7_Click(Index As Integer)
Select Case Index
       Case 0
  a = gk528SetDevice("M20", 1)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
       Case 2
  a = gk528SetDevice("M20", 0)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
       Case 1
       If PDCLWB > 0 And PDCLWB < 5 Then
       csfh = PDCLWB + 1
       sql1 = "UPDATE pldr SET ������Ϣ='ֹͣ',�������='Y' WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "'"
       RD.Open sql1, conn, adOpenStatic, adLockOptimistic
      
       ReDim WriteData(0) As String
       WriteData(0) = csfh
       a = gk528WriteDevice("D121", 1, WriteData)   '��ַ  ����  ��ֵ��
       RWorder = 7
       RWcomm = True
       End If

       Case 5
    Text7 = 0
    Text1(4) = 0
    ReDim WriteData(0) As String
    WriteData(0) = Val(5)
    a = gk528WriteDevice("D121", 1, WriteData)   '��ַ  ����  ��ֵ��
  RWorder = 7
  RWcomm = True
End Select
End Sub

Private Sub Label8_Click(Index As Integer)
Select Case Index
       Case 0
  a = gk528SetDevice("M180", 1)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
       Case 1
  a = gk528SetDevice("M180", 0)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
       Case 2
  a = gk528SetDevice("M190", 1)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
       Case 3
  a = gk528SetDevice("M190", 0)  '��ַ  ��λΪ1 ��λΪ0
  RWorder = 7
  RWcomm = True
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
    
   
   b = ""
   b = MSCONComm(MSComm4)
   
   
   
   If b = "0" Then
      txpd = txpd + 1
   End If
   
      If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
      
   If b <> "0" Then Exit Sub
   Timer4.Enabled = False
   Select Case Order
          Case 0   'read d700-706   ״̬��ȡ
          
                         Ddata(5) = "&H" + Mid(PLCText, 7, 2) + Mid(PLCText, 5, 2) + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2) '*PLC���صļĴ�����ֵ�Ǵӵ��ֽڵ����ֽ����У���������Ҫ��������һ�£�
                         ztdq1(1) = Format(CStr(Val(Ddata(5))), "#0")
                         Label5.Caption = ztdq1(1)
                         
                                                                          
                         Order = 1
                        'read d604--d608  ������   �ϵ����  �������� �������
          Case 1
                         For i = 0 To 2
                         Ddata(i) = "&H" + Mid(PLCText, i * 8 + 7, 2) + Mid(PLCText, i * 8 + 5, 2) + Mid(PLCText, i * 8 + 3, 2) + Mid(PLCText, i * 8 + 1, 2) '*PLC���صļĴ�����ֵ�Ǵӵ��ֽڵ����ֽ����У���������Ҫ��������һ�£�
                         
                         If i = 0 Then
                         ztdq2(i + 1) = Format(CStr(Val(Ddata(0))), "#0")   '''604
                         Text1(5) = Format(ztdq2(i + 1), "#,##0")
                         End If
                         
                         If i = 1 Then
                         ztdq4(1) = Format(CStr(Val(Ddata(1))), "#0")   '''   ��ϴˮ��606
                         Text9 = Format(ztdq4(1), "#,##0")
                         End If
                         
                         If i = 2 Then
                         ztdq4(2) = Format(CStr(Val(Ddata(2))), "#0")   '''������   ˳ʱ����608
                         Text10 = ztdq4(2)
                         End If
                                                 
                         Next i
                                                  
                         
                         Order = 2
                         
                         
        Case 2
        
                          Ddata(1) = "&H" + Mid(PLCText, 11, 2) + Mid(PLCText, 9, 2)
                          sssjxrjc = Val(CStr(Val(Ddata(1))))   ''''д����
                          Text5 = sssjxrjc
                          
                          Ddata(1) = "&H" + Mid(PLCText, 7, 2) + Mid(PLCText, 5, 2)
                          cxsfmjc = Val(CStr(Val(Ddata(1))))   ''''��ϴ���
                         
                          
                          Ddata(0) = "&H" + Mid(PLCText, 3, 2) + Mid(PLCText, 1, 2)
                          PDCLWB = Val(CStr(Val(Ddata(0))))
                          Label6.Caption = PDCLWB
                          Label21.Caption = PDCLWB
                          If PDCLWB = 1 And yxztbl(5) = 1 And yxztbl(6) = 0 Then
                          Label4 = "���������С�������"
                          End If
                          If PDCLWB = 2 Then
                          Label4 = "�ܵ�һ�������С�������"
                          End If
                          If PDCLWB = 3 And yxztbl(6) = 0 Then
                          Label4 = "�ܵ�ˮ��ϴ�С�������"
                          End If
                          If PDCLWB = 4 Then
                          Label4 = "�ܵ����������С�������"
                          End If
                          If PDCLWB = 5 And yxztbl(3) = 0 Then
                          Label4 = "����������ɡ�����"
                          End If
                          
                          Order = 3
                          
         Case 3
         
               Tdata1 = Mid(PLCText, 1, 2)
               Mdata(0) = Val("&H" + Tdata1) ' ת��Ϊʮ����
               Tdata2 = DecToBin(Mdata(0)) '*���ö�����ת������
               Tdata3 = StrReverse(Tdata2)
                      
               Tdata1 = Mid(PLCText, 3, 2)
               Mdata(0) = Val("&H" + Tdata1) ' ת��Ϊʮ����
               Tdata2 = DecToBin(Mdata(0)) '*���ö�����ת������
               Tdata4 = StrReverse(Tdata2)
                                     
               Tdata2 = Tdata3 + Tdata4
               Text3 = Tdata2
                                          
               If Mid(Tdata2, 2, 1) = 1 Then           '''plc����״̬
               yxztbl(7) = 1
               Else
               yxztbl(7) = 0
               End If
                            
               If Mid(Tdata2, 3, 1) = 1 Then          ''''�ܵ�����  ����
               yxztbl(1) = 1
               Label8(0).BackColor = &HFF00&
               Else
               yxztbl(1) = 0
               Label8(0).BackColor = &HC0C0C0
               End If
               
               If Mid(Tdata2, 4, 1) = 1 Then          '''�ܵ�����  ����
               yxztbl(2) = 1
               Label8(2).BackColor = &HFF00&
               Else
               yxztbl(2) = 0
               Label8(2).BackColor = &HC0C0C0
               End If
               
               If Mid(Tdata2, 5, 1) = 1 Then          '''PLC ��ͣ
               Label7(0).BackColor = &HFF00&
               Label7(2).BackColor = &HC0C0C0          '
               yxztbl(3) = 1
               Else
               yxztbl(3) = 0
               Label7(0).BackColor = &HC0C0C0
               Label7(2).BackColor = &HFF00&
               End If

               If Mid(Tdata2, 6, 1) = 1 Then          ''''��̨��ŷ�
               yxztbl(4) = 1
               Label3(3).BackColor = &HFF00&
               Else
               yxztbl(4) = 0
               Label3(3).BackColor = &HC0C0&
               End If

               If Mid(Tdata2, 7, 1) = 1 Then          ''''ԭ�ϵ�ŷ�
               yxztbl(5) = 1
               Label2(6).BackColor = &HFF00&
               Else
               yxztbl(5) = 0
               Label2(6).BackColor = &HC0C0&
               End If
               
               If Mid(Tdata2, 8, 1) = 1 And PDCLWB = 1 Then    ''''ԭ���쳣
               yxztbl(6) = 1
               Else
               yxztbl(6) = 0
               End If
            
               If Mid(Tdata2, 9, 1) = 1 And PDCLWB = 3 Then   '''ˮ�쳣
               yxztbl(8) = 1
               Else
               yxztbl(8) = 0
               End If
               
               If Mid(Tdata2, 10, 1) = 1 Then                '''�������쳣
               SJPD = 1
               Else
               SJPD = 0
               End If
               
               If Mid(Tdata2, 11, 1) = 1 Then                ''''ԭ�ϻ�����  ԭ�Ϸ�����ʱ��
               ysbc = 1
               Else
               ysbc = 0
               End If
               Order = 0
                          
         Case 6, 7, 8  'д �ã���λ
               Order = 0
   End Select

   Timer3.Enabled = True

End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 2
t1 = Val(Format(Text1(4), "#0")) - 500
t2 = Val(Format(Text1(4), "#0")) + 500
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select avg(ƽ��ʱ��) from v_pld_ss_ycjc where Ⱦ��������='" & Text1(2) & "' and �������� between '" & t1 & "' and '" & t2 & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
Text8 = 0
Else
Text8 = Int(Adodc4.Recordset.Fields(0))
End If
End Select
End Sub




Private Sub Timer1_Timer()
On Error Resume Next
       ReDim WriteData(0 To 14) As String  ''''''д�����
       Dim DataW As String    '*���������м䴦�������
       Dim Data10(20) As String
       Dim Buffer(3) As Byte   '*���������м䴦�������


If PDCLWB > 0 Then
Timer1.Enabled = False
End If

 If PDCLWB = 0 Then   ''''''����׼������ʱ
 
 If Text1(0) = "" Or Text1(1) = "" Or Text1(2) = "" Or Text1(3) = "" Or Text1(4) = "" Then Exit Sub
 
       For i = 0 To 5
       Data10(i) = Right("00000000" + Hex(Val(ytsz(i))), 8)
       WriteData(2 * i + 0) = Val("&H" + Right(Data10(i), 4))
       WriteData(2 * i + 1) = Val("&H" + Mid(Data10(i), 1, 4))
       Next
       
       a = gk528WriteDevice("D480", 12, WriteData())
 RWorder = 6
 RWcomm = True
 Timer1.Enabled = False
 xhpdql1 = 1
 xhpdql2 = 1
 xhpdql3 = 1
 xhpdql4 = 1
 xhpdql5 = 1
 xhpdql6 = 1
 xhpdql7 = 1
 xhpdql8 = 1
 Text7 = 0
sql1 = "UPDATE pldr SET ��ʼ����='" & Now & "',����״̬='��ʼ����',������Ϣ='����' WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "' and ��ʼ���� is null"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

End If

If sssjxrjc = 0 Then
Label4.Caption = "��ע��������׼���У���"
End If
End Sub

Private Sub Timer10_Timer()
On Error Resume Next
If PDCLWB = 5 And Val(Text1(5)) > 0 And plcxryc = 1 Then
cll = Format(Val(Format(Text1(5), "#0")) / 1000, "#0.000") ''''''''''������λgת����kg
sql1 = "UPDATE pldr SET ʵ�ʳ���='" & cll & "',��������='" & Now & "',���ͽ���='" & Now & "',�������='Y',����״̬='�������',����ʱ��=Datediff(n,��ʼ����,'" & Now & "'),����ʱ��='" & Text7 & "' WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
plcxryc = 1
wdbl = "0"
If Err.Number = 3709 Then End
End If
End Sub

Private Sub Timer11_Timer()
On Error Resume Next

cll = Format(Val(Format(Text1(5), "#0")) / 1000, "#0.000") ''''''''''������λgת����kg

If PDCLWB = 1 And Val(cll) > 0 Then     ''''����״̬   ����
sql1 = "UPDATE pldr SET ����״̬='����������',ʵ�ʳ���='" & cll & "' WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If PDCLWB = 3 Then        ''''����״̬   ˮ��
sql1 = "UPDATE pldr SET ����״̬='�ܵ�ˮ��ϴ��',ʵ�ʳ���='" & cll & "'  WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "' and ����״̬<>'�ܵ�ˮ��ϴ��'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

If PDCLWB = 4 Then      ''''����״̬   ���������
sql1 = "UPDATE pldr SET ����״̬='�ܵ�����������' WHERE �ϵ����='" & Text1(0) & "' and �����='" & Text1(3) & "' and ��������='" & Text1(1) & "' and ����״̬<>'�ܵ�����������'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If

End Sub

Private Sub Timer12_Timer()
On Error Resume Next
If pdcqplc = 3 Then
pdcqplc = 1
If yxztbl(7) = 0 Then
Unload Me
Me.Show
End If
Else
pdcqplc = pdcqplc + 1
End If
End Sub

Private Sub Timer13_Timer()
If ysbc = 1 Then
Text7 = Val(Text7) + 1
End If
End Sub

Private Sub Timer14_Timer()
If dqdz = 600 Then
Adodc6.RecordSource = "select name, convert(float,size) * (8192.0/1024.0)/1024. from zzpr.dbo.sysfiles"
Adodc6.Refresh
If Not Adodc6.Recordset.EOF Then
Adodc6.Recordset.MoveFirst
Do While Not Adodc6.Recordset.EOF
If Adodc6.Recordset.Fields(0) = "zrrz_log" Then
Text11 = Format(Val(Adodc6.Recordset.Fields(1)), "#0")
If Val(Text11) > 1000 Then
Label4.Caption = "ץ������������־"
End If
End If
Adodc6.Recordset.MoveNext
Loop
End If
dqdz = 1
End If
dqdz = dqdz + 1
End Sub

Private Sub Timer15_Timer()
If txpd > 10 Then
Text6.Text = "ͨ������"
txpd = 0
Else
Text6.Text = "ͨ���쳣"
End If
End Sub

Private Sub VQJC()
On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "SELECT �ϵ����,��������,Ⱦ������,Ⱦ��������,���ϵ�λ,round(��������,4) as ��������,ʵ�ʳ���,�����,��̨,�ܵ�����1,�������,�ܵ�����2,isnull(�ܵ����,0) as �ܵ����,��̨���,���ϵȼ� FROM v_pldr_dx WHERE ����ʱ�� is not null and �Ƿ����='��' and isnull(�������,'')<>'Y' and isnull(������Ϣ,'') not like '%�쳣����%' and �ܵ����<>0 and (��������-isnull(ʵ�ʳ���,0))>0 ORDER BY ����ʱ��,��������,�����"
Adodc2.Refresh

If Adodc2.Recordset.EOF And yxztbl(7) = 1 Then
Timer1.Enabled = False
Timer5.Enabled = True
For i = 0 To 6
Text1(i) = ""
ytsz(i) = ""
Next
Text1(5) = 0
wdbl = "0"
Label4 = "����������ɡ�����"

Else

Adodc2.Recordset.MoveFirst
Adodc3.RecordSource = "SELECT �ϵ����,��������,Ⱦ������,Ⱦ��������,���ϵ�λ,round(��������,4) as ��������,ʵ�ʳ���,�����,��̨,�ܵ�����1,�������,�ܵ�����2,isnull(�ܵ����,0) as �ܵ����,��̨���,���ϵȼ� FROM v_pldr_dx WHERE ����ʱ�� is not null and �Ƿ����='��' and isnull(�������,'')<>'Y' and isnull(������Ϣ,'') not like '%�쳣����%' and �ܵ����<>0 and �ϵ����='" & Adodc2.Recordset.Fields(0) & "' ORDER BY ����ʱ��,��������,�����"
Adodc3.Refresh


If Val(Adodc3.Recordset.RecordCount) > 1 Then
ssgs = 2
Else
ssgs = 1
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ݳ���ת������
If PDCLWB = 0 And yxztbl(7) = 1 Then
Text1(4) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000, "#,##0") ''''������
Text1(0) = Adodc3.Recordset.Fields(0)    '''�ϵ����
Text1(1) = Adodc3.Recordset.Fields(1)    '''��������
Text1(2) = Adodc3.Recordset.Fields(3)       '''''Ⱦ��������
Text1(3) = Adodc3.Recordset.Fields(7)       ''''�����
Text1(5) = 0                             ''''������
Text1(6) = Val(Adodc3.Recordset.Fields(13))  '''��̨���                         '''
Text1(7) = Adodc3.Recordset.Fields(8)  '''��̨                         '''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��������
ytsz(0) = Adodc3.Recordset.Fields(12)    '''�ܵ����      480

If ssgs = 1 Then
If (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 < 3000 Then
ytsz(1) = 100 ''''''�ܵ����  ��һ������ʱ��10��
Else
ytsz(1) = Format(Adodc3.Recordset.Fields(9), "#0")    ''''''�ܵ����  ��һ������ʱ��482
End If
End If

If ssgs = 2 Then
ytsz(1) = 100 ''''''�ܵ����  ��һ������ʱ��10��
End If

If (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 <= 50 Then
ytsz(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 + 200, "#0")
End If

If (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 > 50 And (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 <= 4000 Then
ytsz(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 + (1 - (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 / 6000) * 180, "#0")
End If

If (Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 > 4000 Then
ytsz(2) = Format((Adodc3.Recordset.Fields(5) - Adodc3.Recordset.Fields(6)) * 1000 + 60, "#0") ''''''������  484
End If

If ssgs = 1 Then
If Val(Adodc3.Recordset.Fields(14)) = 1 Then
ytsz(3) = Format(Val(Adodc3.Recordset.Fields(10)), "#0")  ''''''�������   486  ��ϴˮ��
End If

If Val(Adodc3.Recordset.Fields(14)) = 2 Then
ytsz(3) = Format(Val(Adodc3.Recordset.Fields(10)) * 1.5, "#0")  ''''''�������   486  ��ϴˮ��
End If
End If

If ssgs = 2 Then
ytsz(3) = 3000          ''''''�������   486  ��ϴˮ��
End If

If ssgs = 1 Then
ytsz(4) = Format(Adodc3.Recordset.Fields(11), "#0")    ''''''�ܵ���� ��2������ʱ��   488
End If

If ssgs = 2 Then
ytsz(4) = 100   ''''''�ܵ���� ��2������ʱ��   488
End If


ytsz(5) = Val(Adodc3.Recordset.Fields(13))    ''''''��̨���                490
If ytsz(5) >= 1 And ytsz(5) <= 8 Then           '''''��������
sssl1 = sssl1 + 1
End If

If ytsz(5) > 8 Then                             '''��������
sssl2 = sssl2 + 1
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��������
ytsz(6) = Val(Adodc3.Recordset.Fields(0))    ''''''�ϵ����
ytsz(7) = Val(Mid(Adodc3.Recordset.Fields(1), 1, 2))  ''''''������
ytsz(8) = Val(Adodc3.Recordset.Fields(7))    ''''''�������


wdbl = "1"                                '''''''''�ȶ�����
wdbj = "1"                            ''''''''''''''''д����
If sssjxrjc = 0 Then
Label4 = "��������׼���С�����"
End If
Timer1.Enabled = True
End If
End If
End Sub

Private Sub Timer3_Timer()    ''''''''''''''PLC

 If RWcomm = True Then
   Order = RWorder
   RWcomm = False
 End If
  Select Case Order
         Case 0   '��D56
              a = gk528ReadDevice("D700", 2)
         Case 1
              a = gk528ReadDevice("D604", 8)
         Case 2
              a = gk528ReadDevice("D121", 3)
         Case 3
              a = gk528ReadDevice("M16", 10)
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
 
       If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
  
   Timer3.Enabled = True
   RWcomm = False
   Order = 0
 Else
    Timer3.Enabled = False
 End If

End Sub

Private Sub Timer5_Timer()
On Error Resume Next
Label10(0).Caption = sssl1
Label10(1).Caption = sssl2

If sssl1 >= 20 And Adodc2.Recordset.EOF Then
'  A = gk528SetDevice("M180", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
  sssl1 = 0
End If

If sssl2 >= 20 And Adodc2.Recordset.EOF Then
'  A = gk528SetDevice("M190", 1)  '��ַ  ��λΪ1 ��λΪ0
'  RWorder = 7
'  RWcomm = True
  sssl2 = 0
End If

If PDCLWB = 0 And yxztbl(3) = 0 And yxztbl(6) = 0 And yxztbl(7) = 1 And yxztbl(1) = 0 And yxztbl(2) = 0 And Val(Text1(5)) = 0 Then
If plcxryc = 6 Then
Call VQJC
plcxryc = 1
Else
plcxryc = plcxryc + 1
End If
Else
plcxryc = 1
End If
Label10(2).Caption = plcxryc
End Sub


Private Sub Timer6_Timer()
On Error Resume Next
If wdbj = "0" Then
Timer1.Enabled = False
End If

If wdbj = "1" And PDCLWB = 0 Then
Beep 2000, 50
qpys = 3
Timer1.Enabled = True
wdbj = ""
End If
End Sub

Private Sub Timer7_Timer()
On Error Resume Next
If MSComm4.PortOpen = True Then Exit Sub
MSComm4.PortOpen = True
      If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
End Sub

Private Sub Timer8_Timer()
On Error Resume Next
If Text1(0) <> ztdq4(1) And Val(Text1(1)) <> Val(ztdq4(2)) And Text1(3) <> ztdq4(3) Then
Adodc5.RecordSource = "SELECT �ϵ����,��������,Ⱦ������,Ⱦ��������,���ϵ�λ,round(��������,4) as ��������,ʵ�ʳ���,�����,��̨,�ܵ�����1,�������,�ܵ�����2,isnull(�ܵ����,0) as �ܵ����,��̨���,���ϵȼ� FROM v_pldr_dx WHERE �ϵ����='" & ztdq4(1) & "' and left(��������,2) like '" & ztdq4(2) & "'+'%' and �����='" & ztdq4(3) & "'"
Adodc5.Refresh
If Not Adodc5.Recordset.EOF Then
Text1(0) = Adodc2.Recordset.Fields(0)    '''�ϵ����
Text1(1) = Adodc2.Recordset.Fields(1)    '''��������
Text1(2) = Adodc2.Recordset.Fields(3)       '''''Ⱦ��������
Text1(3) = Adodc2.Recordset.Fields(7)       ''''�����
Text1(4) = Format((Adodc2.Recordset.Fields(5) - Adodc2.Recordset.Fields(6)) * 1000, "#,##0") ''''������
Text1(6) = Val(Adodc2.Recordset.Fields(12))  '''��̨���                         '''
Text1(7) = Adodc2.Recordset.Fields(8)  '''��̨
End If
End If
End Sub

Private Sub Timer9_Timer()
On Error Resume Next
If xhpdtx = 5 And yxztbl(7) = 0 Then
Label4 = "ͨѶ���� û�����ӡ�������"
Timer7.Enabled = True
End If

If xhpdtx = 10 Then
yxztbl(7) = 0
xhpdtx = 1
Timer7.Enabled = False
Else
xhpdtx = xhpdtx + 1
End If
End Sub



