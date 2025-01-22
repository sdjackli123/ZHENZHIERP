VERSION 5.00
Begin VB.Form Formr440 
   BackColor       =   &H00C0E0FF&
   Caption         =   "键盘信息"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   12630
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   11055
   End
   Begin VB.Label Label1 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   9240
      TabIndex        =   24
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   6360
      TabIndex        =   23
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   3480
      TabIndex        =   22
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   18
      Left            =   600
      TabIndex        =   21
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   600
      TabIndex        =   20
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   600
      TabIndex        =   19
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   600
      TabIndex        =   18
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   600
      TabIndex        =   17
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   4
      Left            =   3480
      TabIndex        =   16
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   5
      Left            =   3480
      TabIndex        =   15
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   6
      Left            =   3480
      TabIndex        =   14
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   7
      Left            =   3480
      TabIndex        =   13
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   8
      Left            =   6360
      TabIndex        =   12
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   9
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   10
      Left            =   6360
      TabIndex        =   10
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   11
      Left            =   6360
      TabIndex        =   9
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   12
      Left            =   9240
      TabIndex        =   8
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Index           =   13
      Left            =   9240
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   6
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "←删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   3480
      TabIndex        =   4
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   600
      TabIndex        =   3
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   9240
      TabIndex        =   2
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   6360
      TabIndex        =   1
      Top             =   6600
      Width           =   2415
   End
End
Attribute VB_Name = "Formr440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1 = ""
End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case Index
Text1 = Trim(Text1) + Label1(Index).Caption
End Select
End Sub

Private Sub Label2_Click()
If pmbl = 1 Then
Formr331.Text3 = Text1
End If
If pmbl = 2 Then
Formr339.Text3 = Text1
End If
If pmbl = 3 Then
frmLogin.UserName = Text1
End If
If pmbl = 4 Then
frmLogin.Password = Text1
End If
If pmbl = 6 Then
Formr441.Text3 = Text1
End If
If pmbl = 7 Then
Forms511.Text2 = Text1
End If
Unload Me
End Sub

Private Sub Label3_Click()
If Len(Text1) > 0 Then
Text1 = Left(Text1, Len(Text1) - 1)
End If
End Sub

