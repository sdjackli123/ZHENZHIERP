VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy113 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���۸���"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data4 
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ۿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   480
      Width           =   855
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
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
      Top             =   9960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data2 
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy113.frx":0000
      Height          =   390
      Left            =   1560
      TabIndex        =   7
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListField       =   "���"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
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
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȡ��"
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy113.frx":0014
      Height          =   8655
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   15266
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      FormatString    =   "��¼�� "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy113.frx":0028
      Height          =   8655
      Left            =   8040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   15266
      _Version        =   393216
      Cols            =   8
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      FormatString    =   "��¼�� "
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy113.frx":003C
      Height          =   390
      Left            =   4920
      TabIndex        =   8
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      Style           =   2
      ListField       =   "���"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ۿ�"
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
      Left            =   8040
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ƿͻ�"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����ƿͻ�"
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
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Formy113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("ȷ��ȡ���ͻ�������" + DBCombo2.Text, vbYesNo) = vbNo Then Exit Sub
Data4.Database.Execute "delete * from KSBJ where �ͻ�='" & DBCombo2.Text & "'"
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If MsgBox(DBCombo2.Text + "ȷ�����ۿ���" + Text1.Text, vbYesNo) = vbNo Then Exit Sub
Data4.Database.Execute "update KSBJ set �ۿ�='" & Text1.Text & "' where �ͻ�='" & DBCombo2.Text & "'"
Data4.Database.Execute "update KSBJ set ���=format(val(����)*val(�ۿ�),'#0.00') where �ͻ�='" & DBCombo2.Text & "'"
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Command5_Click()
If DBCombo1.Text = "" Or DBCombo2.Text = "" Then
MsgBox ("���ܸ���")
Exit Sub
End If

If Text1.Text = "" Then
MsgBox ("�������ۿ�")
Exit Sub
End If

If MsgBox("ȷ���ͻ����۸�����" + "�ѿͻ�" + DBCombo1.Text + "���Ƶ��ͻ�" + DBCombo2.Text + "�ۿ�Ϊ��" + DBCombo2.Text + Text1.Text, vbYesNo) = vbNo Then Exit Sub
Data4.Database.Execute "insert into KSBJ(���,Ʒ��,���,��λ,����,�ۿ�,���,��ͼ) select ���,Ʒ��,���,��λ,����,'" & Text1.Text & "',���,��ͼ from KSBJ where �ͻ�='" & DBCombo1.Text & "'"
Data4.Database.Execute "update KSBJ set �ͻ�='" & DBCombo2.Text & "',���=format(val(����)*val(�ۿ�),'#0.00') where �ͻ�=null"
Data2.Refresh
Data1.Refresh
End Sub

Private Sub DBCombo1_Change()
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data1.RecordSource = "select * from ksbj where �ͻ�='" & DBCombo1.Text & "' order by ���"
Data1.Refresh
End Sub

Private Sub DBCombo1_Click(Area As Integer)
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data1.RecordSource = "select * from ksbj where �ͻ�='" & DBCombo1.Text & "' order by ���"
Data1.Refresh
End Sub

Private Sub DBCombo2_Change()
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data2.RecordSource = "select * from ksbj where �ͻ�='" & DBCombo2.Text & "' order by ���"
Data2.Refresh
End Sub

Private Sub DBCombo2_Click(Area As Integer)
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data2.RecordSource = "select * from ksbj where �ͻ�='" & DBCombo2.Text & "' order by ���"
Data2.Refresh
End Sub

Private Sub Form_Load()
DBCombo1.Text = ""
DBCombo2.Text = ""
Text1.Text = ""
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"
Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"

Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data3.RecordSource = "select ��� from khzl GROUP BY ���"
Data3.Refresh
End Sub
