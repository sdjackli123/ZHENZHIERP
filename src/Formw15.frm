VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formw15 
   BackColor       =   &H00C0E0FF&
   Caption         =   "��Ʒ���"
   ClientHeight    =   11115
   ClientLeft      =   -435
   ClientTop       =   3810
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ų�ѯ"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ڲ�ѯ"
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ʒ����ѯ"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw15.frx":0000
      Height          =   390
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "Ʒ��"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫ�����"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formw15.frx":0014
      Height          =   7575
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   9
      BackColorFixed  =   8421631
      BackColorBkg    =   50372
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formw15.frx":0028
      Height          =   390
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   688
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "����"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81592321
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81592321
      CurrentDate     =   39177
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ����"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ʒ��"
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
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Formw15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPFH where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data3.Database.Execute "UPDATE CPKC SET ����=-���� "
       Data1.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPRK where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO CPKCZ(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,format(SUM(����),'#0') FROM CPKC GROUP BY ����,���,Ʒ��,���,�ͺ�,��λ"
       Data2.Database.Execute "DELETE * FROM CPKCZ WHERE ����<=0"
       Data2.RecordSource = "SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPKCZ"
       Data2.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If DBCombo1.text = "" Then
MsgBox ("������Ʒ��")
Exit Sub
End If
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPFH where Ʒ��='" & DBCombo1.text & "'"
       Data3.Database.Execute "UPDATE CPKC SET ����=-���� "
       Data1.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPRK where ���<>'00000000' AND Ʒ��='" & DBCombo1.text & "'"
       Data2.Database.Execute "INSERT INTO CPKCZ(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,format(SUM(����),'#0') FROM CPKC GROUP BY ����,���,Ʒ��,���,�ͺ�,��λ"
       Data2.RecordSource = "SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPKCZ where ����<>0"
       Data2.Refresh
End Sub

Private Sub Command5_Click()
Call OutDataToExcel(MSFlexGrid1, 6, "��Ʒ���")
End Sub

Private Sub Command7_Click()
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPFH where ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data3.Database.Execute "UPDATE CPKC SET ����=-���� "
       Data1.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPRK where ���<>'00000000' and ���� between cdate('" & DTPicker1.Value & "') and cdate('" & DTPicker2.Value & "')"
       Data2.Database.Execute "INSERT INTO CPKCZ(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,format(SUM(����),'#0') FROM CPKC GROUP BY ����,���,Ʒ��,���,�ͺ�,��λ"
       Data2.RecordSource = "SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPKCZ where ����<>0"
       Data2.Refresh
End Sub

Private Sub Command8_Click()
If DBCombo2.text = "" Then
MsgBox ("�����뵥��")
Exit Sub
End If
       Data1.Database.Execute "DELETE * FROM CPKC"
       Data1.Database.Execute "DELETE * FROM CPKCZ"
       Data3.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPFH where ����='" & DBCombo2.text & "'"
       Data3.Database.Execute "UPDATE CPKC SET ����=-���� "
       Data1.Database.Execute "INSERT INTO CPKC(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPRK where ����='" & DBCombo2.text & "' AND ���<>'00000000'"
       Data2.Database.Execute "INSERT INTO CPKCZ(����,���,Ʒ��,���,�ͺ�,��λ,����) SELECT ����,���,Ʒ��,���,�ͺ�,��λ,format(SUM(����),'#0') FROM CPKC GROUP BY ����,���,Ʒ��,���,�ͺ�,��λ"
       Data2.RecordSource = "SELECT ����,���,Ʒ��,���,�ͺ�,��λ,���� FROM CPKCZ  where ����<>0"
       Data2.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
DBCombo1.text = ""
DBCombo2.text = ""
DTPicker1.Value = Date - 30
DTPicker2.Value = Date
Data1.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"

Data2.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"

Data3.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"

Data4.DatabaseName = "D:\���ݿ�\htgl\2011\CPCK.MDB"
Data4.RecordSource = "SELECT Ʒ�� FROM CPRK GROUP BY Ʒ��"
Data4.Refresh

Data5.DatabaseName = "D:\���ݿ�\htgl\2011\sczyjhd.mdb"
Data5.RecordSource = "select ����  from SCZY_z group by ����"
Data5.Refresh

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(2) = 1500
MSFlexGrid1.ColWidth(3) = 4500
MSFlexGrid1.ColWidth(4) = 1200
MSFlexGrid1.ColWidth(5) = 1200
MSFlexGrid1.ColWidth(6) = 1500

End Sub
