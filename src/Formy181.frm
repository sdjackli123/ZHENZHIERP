VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formy181 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��������"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form59"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ͬɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ͬɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ɫͬɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ɫͬɫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   855
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy181.frx":0000
      Height          =   330
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "��ɫ"
      Text            =   "DBCombo1"
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy181.frx":0014
      Height          =   330
      Left            =   8400
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "��ɫ"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "Formy181.frx":0028
      Height          =   330
      Left            =   1080
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���Ͽ���"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy181.frx":003C
      Height          =   7575
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy181.frx":0050
      Height          =   7575
      Left            =   7560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   20
      BackColorFixed  =   12632319
      ForeColorSel    =   16744703
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Bindings        =   "Formy181.frx":0064
      Height          =   330
      Left            =   12120
      TabIndex        =   17
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ѡ����ɫ"
      Height          =   375
      Index           =   4
      Left            =   11280
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "������ɫ"
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��׼��ɫ"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   495
      Index           =   13
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Formy181"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public S1, S2 As Integer

Private Sub Command1_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("��ѡ���¼��")
Exit Sub
End If
If MsgBox("ȷ����ɫͬɫ��", vbYesNo) = vbNo Then Exit Sub
If S1 < 1 Or S2 < 1 Then
MsgBox ("ѡ����ɫͬɫ��¼")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("ע��ѡ��˳��")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
Data5.Recordset.Edit
Data5.Recordset.Fields(7) = DBCombo2.Text
Data5.Recordset.Update
Else
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
For i = 1 To k + 1
Data5.Recordset.Edit
Data5.Recordset.Fields(7) = DBCombo2.Text
Data5.Recordset.Update
Data5.Recordset.MoveNext
Next
End If
S1 = 0
S2 = 0
Data5.Refresh
End Sub

Private Sub Command5_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("��ѡ���¼��")
Exit Sub
End If
If MsgBox("ȷ����ɫͬɫ��", vbYesNo) = vbNo Then Exit Sub
If S1 < 1 Or S2 < 1 Then
MsgBox ("ѡ����ɫͬɫ��¼")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("ע��ѡ��˳��")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
Data5.Recordset.Edit
Data5.Recordset.Fields(7) = DBCombo4.Text
Data5.Recordset.Update
Else
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
For i = 1 To k + 1
Data5.Recordset.Edit
Data5.Recordset.Fields(7) = DBCombo4.Text
Data5.Recordset.Update
Data5.Recordset.MoveNext
Next
End If
S1 = 0
S2 = 0
Data5.Refresh
End Sub

Private Sub Command6_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("��ѡ���¼��")
Exit Sub
End If
If MsgBox("ȷ��ɾ����", vbYesNo) = vbNo Then Exit Sub
If S1 < 1 Or S2 < 1 Then
MsgBox ("ѡ��ͬɫ��¼")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("ע��ѡ��˳��")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
Data5.Recordset.Delete
Else
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
For i = 1 To k + 1
Data5.Recordset.Delete
Data5.Recordset.MoveNext
Next
End If
S1 = 0
S2 = 0
Data5.Refresh

End Sub

Private Sub Command7_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("��ѡ���¼��")
Exit Sub
End If
If MsgBox("ȷ������ͬɫ��", vbYesNo) = vbNo Then Exit Sub
If S1 < 1 Or S2 < 1 Then
MsgBox ("ѡ������ͬɫ��¼")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("ע��ѡ��˳��")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
Data5.Recordset.Edit
Data5.Recordset.Fields(8) = DBCombo2.Text
Data5.Recordset.Update
Else
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
For i = 1 To k + 1
Data5.Recordset.Edit
Data5.Recordset.Fields(8) = DBCombo2.Text
Data5.Recordset.Update
Data5.Recordset.MoveNext
Next
End If
S1 = 0
S2 = 0
Data5.Refresh

End Sub

Private Sub Command8_Click()
If Data4.Recordset.EOF Then
MsgBox ("û��Ҫ���ɵ�����")
Exit Sub
End If
If DBCombo2.Text = "" Then
MsgBox ("��ѡ�������ɫ")
Exit Sub
End If
If MsgBox("ȷ�����ɲ�����ɫ" + DBCombo2.Text + "��", vbYesNo) = vbNo Then Exit Sub
Data7.Database.Execute "INSERT INTO dlclb(����,���,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������,���Ͽ���,��λ) SELECT ����,���,��������,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������,���Ͽ���,��λ FROM DLCLB WHERE  ���='" & Text2.Text & "' AND ������ɫ='" & DBCombo1.Text & "' AND ���Ͽ���='" & DBCombo3.Text & "'"
Data7.Database.Execute "UPDATE DLCLB SET ������ɫ='" & DBCombo2.Text & "' WHERE ������ɫ=NULL"
Call Command3_Click
End Sub

Private Sub Command9_Click()
If S1 = 0 Or S2 = 0 Then
MsgBox ("��ѡ���¼��")
Exit Sub
End If
If MsgBox("ȷ������ͬɫ��", vbYesNo) = vbNo Then Exit Sub
If S1 < 1 Or S2 < 1 Then
MsgBox ("ѡ������ͬɫ��¼")
Exit Sub
End If
If S1 > S2 Then
MsgBox ("ע��ѡ��˳��")
Exit Sub
End If
k = S2 - S1
If k = 0 Then
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
Data5.Recordset.Edit
Data5.Recordset.Fields(8) = DBCombo4.Text
Data5.Recordset.Update
Else
Data5.Recordset.MoveFirst
Data5.Recordset.Move S1 - 1
For i = 1 To k + 1
Data5.Recordset.Edit
Data5.Recordset.Fields(8) = DBCombo4.Text
Data5.Recordset.Update
Data5.Recordset.MoveNext
Next
End If
S1 = 0
S2 = 0
Data5.Refresh

End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub

Private Sub Command2_Click()
Data5.RecordSource = "select * from dlclb WHERE  ���='" & Text2.Text & "' AND ������ɫ='" & DBCombo2.Text & "' AND ���Ͽ���='" & DBCombo3.Text & "' order by ��λ,��������,��������"
Data5.Refresh
End Sub

Private Sub Command3_Click()
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select ��ɫ from KSNR WHERE  ���='" & Text2.Text & "' GROUP BY ��ɫ"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select ���Ͽ��� from DLCLB WHERE  ���='" & Text2.Text & "' GROUP BY ���Ͽ���"
Data2.Refresh

Data4.RecordSource = "select * from dlclb WHERE  ���='" & Text2.Text & "' AND ������ɫ='" & DBCombo1.Text & "' AND ���Ͽ���='" & DBCombo3.Text & "' order by ��λ,��������,��������"
Data4.Refresh

Data5.RecordSource = "select * from dlclb WHERE  ���='" & Text2.Text & "' AND ������ɫ='" & DBCombo2.Text & "' AND ���Ͽ���='" & DBCombo3.Text & "' order by ��λ,��������,��������"
Data5.Refresh

End Sub

Private Sub Command4_Click()
Unload Me
End Sub


Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
DBCombo1.Text = ""
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
S1 = 0
S2 = 0
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data6.RecordSource = "SELECT YS FROM YS GROUP BY YS"
Data6.Refresh

Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data8.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(3) = 0
MSFlexGrid1.ColWidth(6) = 0
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.ColWidth(11) = 0

MSFlexGrid2.ColWidth(0) = 200
MSFlexGrid2.ColWidth(1) = 0
MSFlexGrid2.ColWidth(2) = 0
MSFlexGrid2.ColWidth(3) = 0
MSFlexGrid2.ColWidth(6) = 0
MSFlexGrid2.ColWidth(7) = 0
MSFlexGrid2.ColWidth(11) = 0

End Sub

Private Sub Text2_Change()
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data1.RecordSource = "select ��ɫ from KSNR WHERE ���='" & Text2.Text & "' GROUP BY ��ɫ"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data2.RecordSource = "select ���Ͽ��� from DLCLB WHERE  ���='" & Text1.Text & "' GROUP BY ���Ͽ���"
Data2.Refresh

End Sub
