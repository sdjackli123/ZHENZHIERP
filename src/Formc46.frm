VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formc46 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�������ϲ�ѯ"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form46"
   ScaleHeight     =   9990
   ScaleWidth      =   10860
   StartUpPosition =   2  '��Ļ����
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc46.frx":0000
      Height          =   7335
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12938
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formc46.frx":0014
      Height          =   330
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "������"
      BoundColumn     =   "������"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22872065
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   22872065
      CurrentDate     =   39177
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ϳ���"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Formc46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���ݺ�,���� FROM KPD WHERE ���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Else
Data1.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���ݺ�,���� FROM KPD WHERE ���ϳ���='" & DBCombo1.Text & "' AND ���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
End If
Call OutDataToExcel(MSFlexGrid1, 8, "���ϳ���" + DBCombo1.Text)
End Sub

Private Sub Command3_Click()
If DBCombo1.Text = "" Then
Data1.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���ݺ�,����,���ϳ��� FROM KPD WHERE ���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Else
Data1.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���ݺ�,����,���ϳ��� FROM KPD WHERE ���ϳ���='" & DBCombo1.Text & "' AND ���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
End If
End Sub


Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.Value
End Sub

Private Sub Form_Load()
Text4.Text = Date
Text5.Text = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DBCombo1.Text = ""
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data2.RecordSource = "select ct.������  from ct group by ct.������ ORDER BY VAL(CT.������)"
Data2.Refresh

Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\CKGL.MDB"
Data1.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,���ݺ�,����,���ϳ��� FROM KPD WHERE ���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
MSFlexGrid1.ColWidth(0) = 200
End Sub
