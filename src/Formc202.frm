VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formc202 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����Ϣ"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   LinkTopic       =   "Form42"
   ScaleHeight     =   8850
   ScaleWidth      =   10275
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   975
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
      Top             =   7800
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formc202.frx":0000
      Height          =   6375
      Left            =   600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
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
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formc202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
MSFlexGrid2.ColWidth(0) = 200
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 1200
End Sub

Private Sub MSFlexGrid2_dblClick()
If Data2.Recordset.EOF Then Exit Sub
rs = MSFlexGrid2.Row
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1


If khbl = 2 Then
Formc2.DBCombo1(1).Text = Data2.Recordset.Fields(0)
Formc2.DBCombo1(2).Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 3 Then
Form42.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 4 Then
Form41.DBCombo1.Text = Data2.Recordset.Fields(1)
Form41.DBCombo3.Text = Data2.Recordset.Fields(0)
Unload Me
End If

If khbl = 5 Then
Form306.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 6 Then
Form502.DBCombo1.Text = Data2.Recordset.Fields(0)
Form502.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 7 Then
Form91.DBCombo1(1).Text = Data2.Recordset.Fields(1)
Form91.DBCombo1(2).Text = Data2.Recordset.Fields(2)
Form91.DBCombo1(3).Text = Data2.Recordset.Fields(3)
Form91.DBCombo1(4).Text = Data2.Recordset.Fields(4)
Unload Me
End If

If khbl = 8 Then
Form503.DBCombo1.Text = Data2.Recordset.Fields(0)
Form503.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 9 Then
Form307.DBCombo1.Text = Data2.Recordset.Fields(0)
Form307.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 10 Then
Form34.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 12 Then
Form307.DBCombo2.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 13 Then
Form34.DBCombo5.Text = Data2.Recordset.Fields(1)
Unload Me
End If

If khbl = 21 Then
Form16.DBCombo4.Text = Data2.Recordset.Fields(1)
Unload Me
End If

End Sub

Private Sub Text1_Change()
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "SELECT cmb.����,cmb.���,SCZY_ZDH.��ʽ,cmb.��ɫ,cmb.����,cmb.���� FROM SCZY_ZDH,cmb WHERE SCZY_ZDH.����=cmb.���� and instr(cmb.���,'" & Text1.Text & "')>0 order BY cmb.��ɫ,cmb.����"
Data2.Refresh
End Sub

