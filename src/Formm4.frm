VERSION 5.00
Begin VB.Form Formm4 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14475
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "仓库材料库龄"
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
      Left            =   1560
      MouseIcon       =   "Formm4.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   4920
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "库龄查询"
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
      Left            =   1800
      MouseIcon       =   "Formm4.frx":16692
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   2040
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "生产调度"
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
      Left            =   11160
      MouseIcon       =   "Formm4.frx":2CD24
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1800
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "生产干预"
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
      Left            =   11160
      MouseIcon       =   "Formm4.frx":433B6
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2640
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "毛坯库存库龄"
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
      Left            =   10560
      MouseIcon       =   "Formm4.frx":59A48
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   3960
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同履约明细"
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
      Left            =   1920
      MouseIcon       =   "Formm4.frx":700DA
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   600
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "整理产量明细"
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
      Left            =   4800
      MouseIcon       =   "Formm4.frx":8676C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   600
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "订单预警设置"
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
      Index           =   141
      Left            =   1920
      MouseIcon       =   "Formm4.frx":9CDFE
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6360
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "染色预警设置"
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
      Index           =   144
      Left            =   1920
      MouseIcon       =   "Formm4.frx":B3490
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7320
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "染色预警信息"
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
      Index           =   145
      Left            =   9360
      MouseIcon       =   "Formm4.frx":C9B22
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   7320
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "订单预警信息"
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
      Index           =   148
      Left            =   9360
      MouseIcon       =   "Formm4.frx":E01B4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6360
      Width           =   2610
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "染机查询"
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
      Index           =   134
      Left            =   1800
      MouseIcon       =   "Formm4.frx":F6846
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3000
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分存查询"
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
      Left            =   7320
      MouseIcon       =   "Formm4.frx":10CED8
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库盘存"
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
      Left            =   7320
      MouseIcon       =   "Formm4.frx":12356A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库操作"
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
      Left            =   7320
      MouseIcon       =   "Formm4.frx":139BFC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库记录"
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
      Left            =   7320
      MouseIcon       =   "Formm4.frx":15028E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   1
      Left            =   7920
      Picture         =   "Formm4.frx":166920
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   0
      Left            =   7920
      Picture         =   "Formm4.frx":166B34
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分存查询"
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
      Left            =   4440
      MouseIcon       =   "Formm4.frx":166D48
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库盘存"
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
      Left            =   4440
      MouseIcon       =   "Formm4.frx":17D3DA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库操作"
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
      Left            =   4440
      MouseIcon       =   "Formm4.frx":193A6C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "分库记录"
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
      Left            =   4440
      MouseIcon       =   "Formm4.frx":1AA0FE
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   7
      Left            =   5040
      Picture         =   "Formm4.frx":1C0790
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   420
      Index           =   8
      Left            =   5040
      Picture         =   "Formm4.frx":1C09A4
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "Formm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
