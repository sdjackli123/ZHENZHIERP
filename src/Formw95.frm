VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Formw95 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ɨ�����"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data8 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Data Data7 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "װ���ӡ"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�굥ˢ��"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
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
      Top             =   9600
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
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
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8160
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw95.frx":0000
      Height          =   6615
      Left            =   5280
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw95.frx":0014
      Height          =   390
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "xm"
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw95.frx":0028
      Height          =   390
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   688
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "mc"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formw95.frx":003C
      Height          =   4575
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw95.frx":0050
      Height          =   2055
      Left            =   5280
      TabIndex        =   13
      Top             =   7200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   18
      BackColorFixed  =   8421631
      BackColorBkg    =   43176
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "������λ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�굥���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ɨ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   8160
      Width           =   975
   End
End
Attribute VB_Name = "Formw95"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'On Error Resume Next
If Text2.Text = "" Then
MsgBox ("�����뷢���굥���")
Exit Sub
End If
Data1.RecordSource = "select * from zxd where ���='" & Text2.Text & "'"
Data1.Refresh
If Data1.Recordset.EOF Then
MsgBox ("�����ڷ������")
Exit Sub
End If
Data1.Recordset.MoveFirst
DBCombo1.Text = Data1.Recordset.Fields(0)
Data4.Database.Execute "delete * from zxdf"    ''''''
Do While Not Data1.Recordset.EOF
For i = 6 To 14
If Val(Data1.Recordset.Fields(i)) > 0 Then
Data4.Database.Execute "insert into zxdf(�ͻ�,���,���,��ɫ,������,���) VALUES('" & Data1.Recordset.Fields(0) & "','" & Data1.Recordset.Fields(1) & "','" & Data1.Recordset.Fields(2) & "','" & Data1.Recordset.Fields(i - 1) & "','" & Data1.Recordset.Fields(i) & "','" & Data1.Recordset.Fields(17) & "')"
End If
i = i + 2
Next
Data1.Recordset.MoveNext
Loop
Data4.Database.Execute "insert into zxdf(�ͻ�,���,���,��ɫ,������,���) select ������λ,���,�ͺ�,���,sum(����),���ݺ� from lsfh where ���ݺ�='" & Text2.Text & "' group by ������λ,���,�ͺ�,���,���ݺ�"
Data4.Database.Execute "update zxdf set ����='1'"
Data4.Database.Execute "update zxdf set ������='0' where ������=null"
Data4.Database.Execute "update zxdf set ������='0' where ������=null"

Data4.Database.Execute "insert into zxdf(�ͻ�,���,���,��ɫ,���,������,������) select �ͻ�,���,���,��ɫ,���,sum(val(������)),sum(val(������)) from zxdf where ���='" & Text2.Text & "' group by �ͻ�,���,���,��ɫ,���"
Data4.Database.Execute "delete * from zxdf where ����='1'"

Data2.RecordSource = "select ������λ,����,���,Ʒ��,���,�ͺ�,��λ,����,����,����Ա from lsfh where ���ݺ�='" & Text2.Text & "'"
Data2.Refresh

Data5.RecordSource = "select ���,���,��ɫ,������,������ from zxdf order by ���,���,��ɫ"
Data5.Refresh
Call sx(MSFlexGrid3)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If Text2.Text = "" Then
MsgBox ("�������굥���")
Exit Sub
End If
Call fhmxdy(Data7, Data3, Text2.Text)
End Sub

Private Sub Form_Load()
Dim l As Integer
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DBCombo1.Text = ""
DBCombo2.Text = ""
m = ""
Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.mdb"

Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.MDB"

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.MDB"
Data2.RecordSource = "select * from lsfh where ���ݺ�='" & Text2.Text & "' order by ��� desc"
Data2.Refresh

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\CPCK.mdb"

Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"

Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"

Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\cpck.mdb"

Data8.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data8.Refresh

MSFlexGrid1.ColWidth(11) = 1200
MSFlexGrid1.ColWidth(10) = 1200
MSFlexGrid3.ColWidth(0) = 200
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid2.ColWidth(0) = 200

End Sub


Private Sub Label5_dblClick()
Data7.RecordSource = "SELECT * FROM LSFH WHERE ����=cdate('" & Date & "')"
Data7.Refresh
If Not Data7.Recordset.EOF Then
Data7.RecordSource = "select max(mid(������,7)) from lsfh where ����=cdate('" & Date & "')"
Data7.Refresh
If Len(Data7.Recordset.Fields(0) + 1) < 2 Then
Text3.Text = "C" + Format(Date, "mmdd") + "-" + "0" + Trim(Data7.Recordset.Fields(0) + 1)
Else
Text3.Text = "C" + Format(Date, "mmdd") + "-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Else
Text3.Text = "C" + Format(Date, "mmdd") + "-" + "01"
End If
End Sub

Private Sub MSFlexGrid1_dblClick()
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
rs = MSFlexGrid1.Row
rc = MSFlexGrid1.Col
Data2.Recordset.Move rs - 1
If rc = 1 Then
Data2.Recordset.Delete
Data2.Refresh
End If
End Sub

Private Sub MSFlexGrid3_DBLClick()
''���,���,��ɫ
If Data5.Recordset.EOF Then Exit Sub
rs = MSFlexGrid3.Row
Data5.Recordset.MoveFirst
Data5.Recordset.Move rs - 1
Data4.Database.Execute "delete * from lscx"
Data4.Database.Execute "insert into lscx(����,���,Ʒ��,���,�ͺ�,��λ,����,����,��ע) select ����,���,Ʒ��,���,�ͺ�,��λ,����,����,��ע from lsrk where ���='" & Data5.Recordset.Fields(0) & "' and ���='" & Data5.Recordset.Fields(2) & "' and �ͺ�='" & Data5.Recordset.Fields(1) & "'"
Data4.Database.Execute "insert into lscx(����,���,Ʒ��,���,�ͺ�,��λ,����,����,��ע) select ����,���,Ʒ��,���,�ͺ�,��λ,-����,����,��ע from lsfh where ���='" & Data5.Recordset.Fields(0) & "' and ���='" & Data5.Recordset.Fields(2) & "' and �ͺ�='" & Data5.Recordset.Fields(1) & "'"
Data4.Database.Execute "update lscx set ����='1'"
Data4.Database.Execute "insert into lscx(����,���,Ʒ��,���,�ͺ�,��λ,����,����,��ע) select ����,���,Ʒ��,���,�ͺ�,��λ,sum(����),����,��ע from lscx group by ����,���,Ʒ��,���,�ͺ�,��λ,����,��ע"
Data4.Database.Execute "delete * from lscx where ����='1' or ����<=0"
Data6.RecordSource = "select ����,���,Ʒ��,���,�ͺ�,��λ,����,����,��ע from lscx"
Data6.Refresh
End Sub

Private Sub Text1_Change()
If DBCombo1.Text = "" Then Exit Sub

If InStr(Text1.Text, "J") > 0 Then
m = Left(Text1.Text, Len(Text1.Text) - 1)

If Len(m) = 9 Then

If Text3.Text = "" Then
MsgBox ("���������")
Exit Sub
End If

Data4.RecordSource = "SELECT * FROM LSRK"
Data4.Refresh

Data4.Recordset.FindFirst "����='" & m & "'"
If Data4.Recordset.NoMatch Then
Label2.Caption = "�����ڴ�����"
Text1.Text = ""
Timer1.Enabled = True
Exit Sub
Else
Data6.RecordSource = "SELECT * FROM LSFH WHERE ����='" & m & "'"
Data6.Refresh
If Data6.Recordset.EOF Then

l = 1
Data3.RecordSource = "SELECT ��� FROM LSFH WHERE ���ݺ�='" & Text2.Text & "' ORDER BY ��� DESC"
Data3.Refresh
If Data3.Recordset.EOF Then
l = 1
Else
l = Data3.Recordset.Fields(0) + 1
End If
Data5.Database.Execute "INSERT INTO lsfh(����,����,���,Ʒ��,���,�ͺ�,��λ,����,��ע,����,���,������λ,����Ա,���ݺ�,������) select ����,����,���,Ʒ��,���,�ͺ�,��λ,����,��ע,����,'" & l & "','" & DBCombo1.Text & "','" & DBCombo2.Text & "','" & Text2.Text & "','" & Text3.Text & "' from lsrk where ����='" & m & "'"
End If
Data2.RecordSource = "select ������λ,����,���,Ʒ��,���,�ͺ�,��λ,����,����,������ as ���,����Ա from lsfh where ���ݺ�='" & Text2.Text & "'"
Data2.Refresh
Text1.Text = ""
Text1.SetFocus
End If

Else
Text1.Text = ""
Text1.SetFocus

End If
End If

End Sub


Private Sub sx(MSF As MSFlexGrid)

    Dim i     As Integer
      With MSF
                 .AllowBigSelection = True           '   ����������ʽ
                 .FillStyle = flexFillRepeat
                For i = 1 To .Rows - 1
                        .Row = i:       .Col = .FixedCols
                        .ColSel = .Cols() - .FixedCols - 1
                         If Val(MSF.TextMatrix(i, 4)) < Val(MSF.TextMatrix(i, 5)) Then
                              .CellBackColor = vbGreen           '��ɫ
                        Else
                              .CellBackColor = vbBlack      ' ��ɫ
                        End If
                Next i
        End With
End Sub

