VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy51 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���ɲɹ���"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form30"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1920
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ�����ȷ"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ�����ȷ"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ�����"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ����¼"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1111 
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Text            =   "Text1111"
      Top             =   5280
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ����ӡ"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ�����ͳһ"
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������Ϣ"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy51.frx":0000
      Height          =   330
      Left            =   8760
      TabIndex        =   21
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "MC"
      Text            =   "DBCombo2"
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����Ϣ"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�鿴�ɹ���"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ɹ���Ϣ"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�鿴���ϱ�"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�������"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy51.frx":0014
      Height          =   1935
      Left            =   3600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "��¼�� "
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy51.frx":0028
      Height          =   5535
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "��¼�� "
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formy51.frx":003C
      Height          =   5535
      Left            =   7560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3960
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "��¼�� "
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39883
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "ѡ�����"
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
      Left            =   7560
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����·�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   19
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9600
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "ˢ��"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "ѡ�񵥺�"
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
      Index           =   1
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
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
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
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
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Formy51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, M1, M2, M3, M4, M5 As String: Public c, r, S1, S2 As Integer
Private Sub Command1_Click()
Data2.RecordSource = "SELECT DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������,SUM(DHCLB.��������) AS ������ FROM DHCLB WHERE DHCLB.����='" & DBCombo1.Text & "' GROUP BY DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
End Sub

Private Sub Command10_Click()
On Error Resume Next
Data2.Recordset.MoveFirst
Data2.Recordset.Move S1 - 1
p = S2 - S1 + 1
For i = 1 To p
Data2.Recordset.Delete
Data2.Recordset.MoveNext
Next
Data2.Refresh
End Sub

Private Sub Command11_Click()
If MsgBox("ȷ���ɹ������� ���ţ�" + DBCombo1.Text, vbYesNo) = vbNo Then Exit Sub
Data2.Database.Execute "UPDATE SCZY_ZDH SET ���='��' WHERE ����='" & DBCombo1.Text & "'"
Data1.Refresh
End Sub

Private Sub Command12_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1='Y' WHERE ����='" & DBCombo1.Text & "'"
End Sub

Private Sub Command13_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1='N' WHERE ����='" & DBCombo1.Text & "'"
End Sub

Private Sub Command14_Click()
Data2.Database.Execute "UPDATE SCZY_ZDH SET B1=''  WHERE B1=NULL AND INSTR(SCZY_ZDH.����,'L')>0 AND (���=NULL OR ���<>'��') AND SCZY_ZDH.���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Data1.Recordset.MoveFirst
p = 1
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields(24) = "Y" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbGreen
End If

If Data1.Recordset.Fields(24) = "N" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbRed
End If

If Data1.Recordset.Fields(24) = "" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbCyan
End If

Data1.Recordset.MoveNext
p = p + 1
Loop

End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.Database.Execute "DELETE * FROM CKGL"

Data3.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�������) IN'd:\���ݿ�\\htgl\2011\SCZYJHD.MDB' SELECT ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,SUM(����) FROM CKGL WHERE CKGL.����='" & DBCombo1.Text & "' GROUP BY ����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,���� "
Data1.Database.Execute "UPDATE CKGL SET LX=CK,�ɹ�����=0 WHERE LX=NULL"
Data1.Database.Execute "INSERT INTO CKGL(����,���Ͽ���,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,�ɹ�����) SELECT CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,SUM(CGCLB.��������) AS �ɹ����� FROM CGCLB WHERE CGCLB.����='" & DBCombo1.Text & "' GROUP BY CGCLB.����,CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������"
Data1.Database.Execute "UPDATE CKGL SET LX=CK,�������=0 WHERE LX=NULL"
Data2.RecordSource = "SELECT CKGL.����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.������ɫ,SUM(CKGL.�ɹ�����) AS �ɹ���,SUM(CKGL.�������) AS ����� FROM CKGL WHERE  CKGL.����='" & DBCombo1.Text & "' GROUP BY CKGL.����,CKGL.���Ͽ���,CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.������ɫ"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
Call SX2(Data2, MSFlexGrid2, 8)
End Sub

Private Sub Command3_Click()
Data3.Database.Execute "DELETE * FROM CLRCZZ"
Data3.Database.Execute "DELETE * FROM CLRCZZHZ"
Data3.Database.Execute "INSERT INTO CLRCZZ(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.��ɫ,CKGL.����,CKGL.����,CKGL.����,CKGL.���� from ckgl WHERE CKGL.���='�����' AND CKGL.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data3.Database.Execute "UPDATE CLRCZZ SET ���='���' where ���=NULL"
Data3.Database.Execute "INSERT INTO CLRCZZ(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select CKBL.��������,CKBL.���Ϲ��,CKBL.���ϵ�λ,CKBL.��ɫ,CKBL.����,CKBL.����,CKBL.����,CKBL.���� from ckBL WHERE CKBL.���='�����' AND CKBL.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data3.Database.Execute "UPDATE CLRCZZ SET ���='����',����=-���� WHERE ���=NULL"
Data3.Database.Execute "INSERT INTO CLRCZZHZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����) SELECT CLRCZZ.����,CLRCZZ.��������,CLRCZZ.���Ϲ��,CLRCZZ.���ϵ�λ,CLRCZZ.��ɫ,CLRCZZ.����,SUM(CLRCZZ.����) AS L,AVG(CLRCZZ.����) AS D FROM CLRCZZ GROUP BY CLRCZZ.����,CLRCZZ.��������,CLRCZZ.���Ϲ��,CLRCZZ.���ϵ�λ,CLRCZZ.��ɫ,CLRCZZ.����"
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����>0"
Data4.Refresh
End Sub

Private Sub Command4_Click()
Data2.RecordSource = "SELECT CGCLB.���Ͽ���,CGCLB.��������,CGCLB.���Ϲ��,CGCLB.���ϵ�λ,CGCLB.������ɫ,CGCLB.��������,CGCLB.�������� AS �ɹ��� FROM CGCLB WHERE CGCLB.����='" & DBCombo1.Text & "' AND CGCLB.��������>0 ORDER BY ���Ͽ���,��������,CGCLB.���Ϲ��,������ɫ"
Data2.Refresh
Call SX2(Data2, MSFlexGrid2, 7)
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����='" & DBCombo2.Text & "' AND CLRCZZHZ.����>0"
Data4.Refresh
End Sub

Private Sub Command7_Click()
On Error Resume Next
Formy52.DBCombo1(12).Text = Data4.Recordset.Fields(7)
Formy52.DBCombo1(3).Text = Data4.Recordset.Fields(0)
Formy52.DBCombo2.Text = Data4.Recordset.Fields(3)
Formy52.DBCombo1(1).Text = DBCombo1.Text
Formy52.Text2.Text = DBCombo1.Text
End Sub

Private Sub Command8_Click()
l = Format(Date, "YYMMDD")
Data2.Database.Execute "UPDATE CGCLB SET ��������='" & l & "' WHERE ����='" & DBCombo1.Text & "'"
Data2.Refresh
End Sub

Private Sub Command9_Click()
On Error Resume Next
If Data2.Recordset.EOF Then
MsgBox ("�޼�¼���ܴ�ӡ��")
Exit Sub
End If
Call MXOutDataToExcel(MSFlexGrid2, "���ţ� " + DBCombo1.Text + "��Լ�ţ�" + Data1.Recordset.Fields(8) + "  �ɹ���")
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

Private Sub DTPicker3_Change()
Text3.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text3.Text = Month(DTPicker3.Value)
End Sub


Private Sub Form_Load()
DTPicker3.Value = Date
Text3.Text = Month(DTPicker1.Value)
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select
Text4.Text = Date - 15
Text5.Text = Date
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
DBCombo1.Text = ""
DBCombo2.Text = ""
Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data1.RecordSource = "SELECT * FROM SCZY_ZDH WHERE INSTR(SCZY_ZDH.����,'L')>0 AND (���=NULL OR ���<>'��') AND SCZY_ZDH.���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.MDB"
Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.MDB"
Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.MDB"
Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.MDB"
Data5.RecordSource = "SELECT KL.MC FROM KL GROUP BY KL.MC"
Data5.Refresh
MSFlexGrid1.ColWidth(8) = 1500
End Sub

Private Sub Label2_Click()
Data1.RecordSource = "SELECT * FROM SCZY_ZDH WHERE INSTR(SCZY_ZDH.����,'L')>0 AND (���=NULL OR ���<>'��') AND SCZY_ZDH.���� BETWEEN CDATE('" & Text4.Text & "') AND CDATE('" & Text5.Text & "')"
Data1.Refresh
End Sub

Private Sub MSFlexGrid1_dblClick()
rs = MSFlexGrid1.Row
If Data1.Recordset.EOF Then
DBCombo1.Text = ""
Exit Sub
End If

Data1.Recordset.MoveFirst
Data1.Recordset.Move rs - 1
DBCombo1.Text = Data1.Recordset.Fields(7)
End Sub

Private Sub MSFlexGrid2_Click()
On Error Resume Next
rs = MSFlexGrid2.Row
'If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
DBCombo2.Text = Data2.Recordset.Fields(0)
Data4.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����='" & Data2.Recordset.Fields(0) & "' AND CLRCZZHZ.��������='" & Data2.Recordset.Fields(1) & "' AND ��ɫ='" & Data2.Recordset.Fields(4) & "' AND CLRCZZHZ.����>0"
Data4.Refresh
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
S1 = MSFlexGrid2.RowSel
End Sub

Private Sub MSFlexGrid2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
S2 = MSFlexGrid2.RowSel
End Sub


Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
rs = MSFlexGrid3.Row
Data4.Recordset.MoveFirst
Data4.Recordset.Move rs - 1
Formy52.DBCombo1(12).Text = Data4.Recordset.Fields(7)
Formy52.DBCombo1(3).Text = Data4.Recordset.Fields(0)
Formy52.DBCombo2.Text = Data4.Recordset.Fields(3)
Formy52.DBCombo1(1).Text = DBCombo1.Text
End Sub

Private Sub Text3_Change()
Select Case Text3.Text
       Case 1
K1 = Format(Date, "YYYY") + "-" + "01" + "-01"
K2 = Format(Date, "YYYY") + "-" + "01" + "-31"
       Case 2
If Val(Format(Date, "YYYY")) / 4 = Int(Val(Format(Date, "YYYY")) / 4) Then
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-29"
Else
K1 = Format(Date, "YYYY") + "-" + "02" + "-01"
K2 = Format(Date, "YYYY") + "-" + "02" + "-28"
End If
       Case 3
K1 = Format(Date, "YYYY") + "-" + "03" + "-01"
K2 = Format(Date, "YYYY") + "-" + "03" + "-31"
       Case 4
K1 = Format(Date, "YYYY") + "-" + "04" + "-01"
K2 = Format(Date, "YYYY") + "-" + "04" + "-30"
       Case 5
K1 = Format(Date, "YYYY") + "-" + "05" + "-01"
K2 = Format(Date, "YYYY") + "-" + "05" + "-31"
       Case 6
K1 = Format(Date, "YYYY") + "-" + "06" + "-01"
K2 = Format(Date, "YYYY") + "-" + "06" + "-30"
       Case 7
K1 = Format(Date, "YYYY") + "-" + "07" + "-01"
K2 = Format(Date, "YYYY") + "-" + "07" + "-31"
       Case 8
K1 = Format(Date, "YYYY") + "-" + "08" + "-01"
K2 = Format(Date, "YYYY") + "-" + "08" + "-30"
       Case 9
K1 = Format(Date, "YYYY") + "-" + "09" + "-01"
K2 = Format(Date, "YYYY") + "-" + "09" + "-31"
       Case 10
K1 = Format(Date, "YYYY") + "-" + "10" + "-01"
K2 = Format(Date, "YYYY") + "-" + "10" + "-30"
       Case 11
K1 = Format(Date, "YYYY") + "-" + "11" + "-01"
K2 = Format(Date, "YYYY") + "-" + "11" + "-31"
       Case 12
K1 = Format(Date, "YYYY") + "-" + "12" + "-01"
K2 = Format(Date, "YYYY") + "-" + "12" + "-30"
End Select

End Sub
Private Sub MSFlexGrid2_dblClick()
With MSFlexGrid2
    c = .Col: r = .Row
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSFlexGrid2_dblClick
End If
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid2.SetFocus

    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid2.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid2.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data2.Recordset.MoveFirst
Data2.Recordset.Move r - 1
Data2.Recordset.Edit
Data2.Recordset.Fields(c - 1) = Text1111.Text
Data2.Recordset.Update
Text1111.Visible = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Data1.Refresh
Data1.Recordset.MoveFirst
p = 1
Do While Not Data1.Recordset.EOF

If Data1.Recordset.Fields(24) = "Y" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbGreen
End If

If Data1.Recordset.Fields(24) = Null Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbRed
End If

If Data1.Recordset.Fields(24) = "N" Then
    MSFlexGrid1.Row = p
    MSFlexGrid1.Col = 7 + 1
    MSFlexGrid1.CellBackColor = vbCyan
End If

Data1.Recordset.MoveNext
p = p + 1
Loop

End Sub

