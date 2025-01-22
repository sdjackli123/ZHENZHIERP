VERSION 5.00
Begin VB.Form formy17 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "formy17.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8640
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "联系电话：15965780414；15054498082"
         Height          =   495
         Left            =   3360
         TabIndex        =   9
         Top             =   3120
         Width           =   4815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "（现为浙江理工大学）"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "软件名称："
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "主设计：柳邦军"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "formy17.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2175
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00C0E0FF&
         Caption         =   "本软件所有权归高密市富源软件有限公司所有"
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00C0E0FF&
         Caption         =   "警告:严禁私自复制！因私自复制使用造成的后果自行承担！！"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   4965
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "版本：2009NO1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   3
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "ll"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "设计组毕业于：浙江丝绸工学院染整工程分院"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   5100
      End
   End
End
Attribute VB_Name = "formy17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
Formy6.Show
End Sub

Private Sub Formy_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
   
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
