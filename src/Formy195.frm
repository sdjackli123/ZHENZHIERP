VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy195 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��Э��ѯ"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   LinkTopic       =   "Form19"
   ScaleHeight     =   9375
   ScaleWidth      =   11205
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��λ��ѯ"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ų�ѯ"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "׼��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy195.frx":0000
      Height          =   7575
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy195.frx":0014
      Height          =   330
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ͻ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Formy195"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call WXCX(MSFlexGrid1, "��Э��ѯ")
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Data1.Database.Execute "delete * from wxcx"
Data1.Database.Execute "insert into wxcx(����,���,��ɫ,���,λ��,���,��λ,����) select ����,���,��ɫ,���,λ��,���,��λ,���� from wxjl"
Data1.Database.Execute "insert into wxcx(����,���,��ɫ,���,λ��,���,��λ,����) select ����,���,��ɫ,���,λ��,���,��λ,-val(����) from wxrk"
Data1.Database.Execute "update wxcx set ����='1'"
Data1.Database.Execute "insert into wxcx (����,���,��ɫ,���,λ��,���,��λ,����) select ����,���,��ɫ,���,λ��,���,��λ,sum(val(����)) as ���� from wxcx group by ����,���,��ɫ,���,λ��,���,��λ"
Data1.Database.Execute "delete * from wxcx where ����='1'"
Command2.Enabled = True
End Sub
Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Data2.RecordSource = "select ����,���,��ɫ,���,λ��,���,��λ,���� as ���� from wxcx where ����='" & Text2.Text & "' and  val(����)>0 order by ���,��ɫ,���,λ��"
Data2.Refresh
End Sub

Private Sub Command5_Click()
Data2.RecordSource = "select ����,���,��ɫ,���,λ��,���,��λ,���� as ���� from wxcx where ��λ='" & DBCombo1.Text & "' and  val(����)>0 order by ����,���,��ɫ,���,λ��"
Data2.Refresh
End Sub

Private Sub Form_Load()
Text2.Text = ""
DBCombo1.Text = ""
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\scjd.mdb"
Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select ��� from WXZL group by ���"
Data3.Refresh
MSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid1.ColWidth(i) = 1200
Next

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 0
       khbl = 12
Formy202.Show
End Select
End Sub
