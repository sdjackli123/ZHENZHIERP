VERSION 5.00
Begin VB.Form Forma170 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��״̬ѡ��"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   6960
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ѡ��"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   7200
      ItemData        =   "Forma170.frx":0000
      Left            =   600
      List            =   "Forma170.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "Forma170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
List3.Clear
List3.AddItem "�ƻ�"
List3.AddItem "�ѱ�����Ⱦ"
List3.AddItem "Ԥ����Ⱦɫ"
List3.AddItem "Ⱦɫ��"
List3.AddItem "Ⱦɫ���"
List3.AddItem "�����ĥë"
List3.AddItem "�����ӡ��"
List3.AddItem "ĥë"
List3.AddItem "���Ͱ�װ"
List3.AddItem "�������"
End Sub

Private Sub Command2_Click()
On Error Resume Next
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next
End Sub

Private Sub Command3_Click()
dxcx = ""
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
dxcx = dxcx + List3.List(i) + "-"
End If
Next
Unload Me
End Sub

Private Sub Command4_Click()
dxcx = ""
Unload Me
End Sub
