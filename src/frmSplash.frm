VERSION 5.00
Begin VB.Form form7 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label3 
         Caption         =   "����Ϊ�㽭����ѧ��"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "������ƣ�"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "����ƣ������"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2175
      End
      Begin VB.Label lblCopyright 
         Caption         =   "���������Ȩ�����������ҵ����������޹�˾����"
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "����:�Ͻ�˽�Ը��ƣ���˽�Ը���ʹ����ɵĺ�����ге�����"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6765
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�汾��2007NO1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   3
         Top             =   2160
         Width           =   1710
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "ll"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         TabIndex        =   5
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "������ҵ�ڣ��㽭˿��ѧԺȾ�����̷�Ժ"
         BeginProperty Font 
            Name            =   "����"
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
Attribute VB_Name = "form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
form6.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
   
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
