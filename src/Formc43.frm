VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Formc43 
   BackColor       =   &H00C0E0FF&
   Caption         =   "���䱸��"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form43"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Top             =   10320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӱ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1111 
      Height          =   495
      Left            =   1560
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   6960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
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
      Top             =   10200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�������ϱ�"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   10440
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�µ��ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   11520
      Top             =   5760
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ɾ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���Ϸ�ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      TabIndex        =   18
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option5 
         BackColor       =   &H0000C0C0&
         Caption         =   "�ɹ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000C0C0&
         Caption         =   "��汸��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   10200
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   10200
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   7
      Left            =   13080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   8
      Left            =   13080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   13080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   13080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   11
      Left            =   13080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�������"
      Height          =   495
      Left            =   11880
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option6 
         BackColor       =   &H0000C0C0&
         Caption         =   "����"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H0000C0C0&
         Caption         =   "ǿ��"
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   5040
      TabIndex        =   22
      Top             =   6360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Formc43.frx":0000
      Height          =   1335
      Left            =   3360
      TabIndex        =   29
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   17
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formc43.frx":0015
      Height          =   2655
      Left            =   240
      TabIndex        =   31
      Top             =   3840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   30
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formc43.frx":0029
      Height          =   2895
      Left            =   240
      TabIndex        =   32
      Top             =   6840
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   5106
      _Version        =   393216
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Left            =   1320
      TabIndex        =   33
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Bindings        =   "Formc43.frx":003D
      Height          =   330
      Left            =   10080
      TabIndex        =   34
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "������"
      Text            =   "DBCombo3"
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   96010241
      CurrentDate     =   39883
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   56
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   96010241
      CurrentDate     =   39557
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   57
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   96010241
      CurrentDate     =   39557
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formc43.frx":0051
      Height          =   1815
      Left            =   240
      TabIndex        =   60
      Top             =   1920
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   17
      BackColorFixed  =   8421631
      BackColorBkg    =   39835
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ʼ���ڣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   59
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "�������ڣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   58
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����Ϣ"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����¼"
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   53
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   52
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ϳ��䣺"
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
      Index           =   6
      Left            =   9000
      TabIndex        =   51
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   0
      Left            =   9120
      TabIndex        =   50
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   49
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�����·�"
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
      Index           =   3
      Left            =   240
      TabIndex        =   48
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Left            =   3360
      TabIndex        =   47
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   46
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   45
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ɫ"
      Height          =   375
      Index           =   3
      Left            =   9120
      TabIndex        =   44
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "��λ"
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   43
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   42
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   4
      Left            =   9120
      TabIndex        =   41
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   5
      Left            =   12000
      TabIndex        =   40
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "�ϼƽ��"
      Height          =   375
      Index           =   5
      Left            =   12000
      TabIndex        =   39
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      Height          =   375
      Index           =   6
      Left            =   12000
      TabIndex        =   38
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   375
      Index           =   7
      Left            =   12000
      TabIndex        =   37
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "��ע"
      Height          =   375
      Index           =   6
      Left            =   12000
      TabIndex        =   36
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "Formc43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, KB As String
Public c, r, BAR As Integer
Private Sub Combo1_Change()
End Sub

Private Sub Combo1_Click()
End Sub

Private Sub Command10_Click()
If Data6.Recordset.EOF Then
MsgBox ("�˵��ݺ����޼�¼�����ܴ�ӡ��")
Exit Sub
End If
BAR = 1
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command11_Click()
On Error Resume Next
If MsgBox("ȷ��ɾ����ɾ�������ָܻ���", vbYesNo) = vbNo Then Exit Sub
Data6.Recordset.Delete
Data6.Refresh
Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 11
Text1(i).Text = ""
Next

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Option6.Value = True Then
If DBCombo5.Text = "" Then
MsgBox ("���ϳ���")
Exit Sub
End If

If Data3.Recordset.EOF Then
MsgBox ("�޿�棬���ܽ��У�")
Exit Sub
End If

If Text3.Text = "" Then
MsgBox ("�޳����������ܽ��У�")
Exit Sub
End If
End If


If Option7.Value = True Then
If DBCombo5.Text = "" Then
MsgBox ("���ϳ���")
Exit Sub
End If

If Data3.Recordset.EOF Then
MsgBox ("�޿�棬���ܽ��У�")
Exit Sub
End If

If Text3.Text = "" Then
MsgBox ("�޳����������ܽ��У�")
Exit Sub
End If


End If




Data6.Recordset.AddNew
Data6.Recordset.Fields(0) = ""
Data6.Recordset.Fields(1) = DBCombo4.Text
Data6.Recordset.Fields(2) = DBCombo1.Text
Data6.Recordset.Fields(3) = Text1(0).Text
Data6.Recordset.Fields(4) = Text1(1).Text
Data6.Recordset.Fields(5) = Text1(2).Text
Data6.Recordset.Fields(6) = Text1(3).Text
Data6.Recordset.Fields(7) = Text1(4).Text
Data6.Recordset.Fields(8) = Text1(5).Text
Data6.Recordset.Fields(9) = Text1(6).Text
Data6.Recordset.Fields(10) = Text1(7).Text
Data6.Recordset.Fields(11) = DBCombo3.Text
Data6.Recordset.Fields(12) = Text1(9).Text
Data6.Recordset.Fields(13) = Text1(10).Text
Data6.Recordset.Fields(14) = Text2.Text
Data6.Recordset.Fields(15) = Text1(11).Text
Data6.Recordset.Fields(16) = Text1(8).Text
Data6.Recordset.Fields(17) = KB
Data6.Recordset.Fields(18) = DBCombo5.Text
Data6.Recordset.Fields(19) = "δ"
Data6.Recordset.Fields(20) = "δ"
Data6.Recordset.Update
Data6.Refresh
Data8.Refresh

If Data6.Recordset.RecordCount = 8 Then
If MsgBox("�Ƿ��ӡ�����ݣ�", vbYesNo) = vbNo Then
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If
Else       '''''''''''''''''''''''''''
Call Command10_Click
End If
End If
Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 11
Text1(i).Text = ""
Next
End Sub




Private Sub Command4_Click()
If MsgBox("ȷ���޸���", vbYesNo) = vbNo Then Exit Sub
Data6.Recordset.Edit
Data6.Recordset.Fields(0) = ""
Data6.Recordset.Fields(1) = DBCombo4.Text
Data6.Recordset.Fields(3) = Text1(0).Text
Data6.Recordset.Fields(4) = Text1(1).Text
Data6.Recordset.Fields(5) = Text1(2).Text
Data6.Recordset.Fields(6) = Text1(3).Text
Data6.Recordset.Fields(7) = Text1(4).Text
Data6.Recordset.Fields(8) = Text1(5).Text
Data6.Recordset.Fields(9) = Text1(6).Text
Data6.Recordset.Fields(10) = Text1(7).Text
Data6.Recordset.Fields(12) = Text1(9).Text
Data6.Recordset.Fields(13) = Text1(10).Text
Data6.Recordset.Fields(14) = Text2.Text
Data6.Recordset.Fields(15) = Text1(11).Text
Data6.Recordset.Fields(16) = Text1(8).Text
Data6.Recordset.Fields(17) = KB
Data6.Recordset.Fields(18) = DBCombo5.Text
Data6.Recordset.Update
Data6.Refresh
Data8.Refresh

Call Command7_Click
Option1.Value = False
Option5.Value = False
For i = 0 To 11
Text1(i).Text = ""
Next

End Sub

Private Sub Command7_Click()
Data1.Database.Execute "DELETE * FROM JHCK"
Data1.Database.Execute "INSERT INTO JHCK(����,��������,���Ϲ��,���ϵ�λ,������ɫ,�ƻ���) SELECT DHCLB.����,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,SUM(DHCLB.��������) AS �ƻ��� FROM DHCLB WHERE DHCLB.����='" & DBCombo4.Text & "' GROUP BY DHCLB.����,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ"
Data3.Database.Execute "INSERT INTO JHCK(����,��������,���Ϲ��,���ϵ�λ,������ɫ,ʵ����) IN 'd:\���ݿ�\htgl\2011\SCZYJHD.MDB' SELECT KPD.����,KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ,SUM(KPD.����) AS ʵ���� FROM KPD WHERE KPD.����='" & DBCombo4.Text & "' GROUP BY KPD.����,KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ"
Data1.Database.Execute "UPDATE JHCK SET �ƻ���=0 WHERE �ƻ���=null"
Data1.Database.Execute "UPDATE JHCK SET ʵ����=0 WHERE ʵ����=null"
Data1.Database.Execute "UPDATE JHCK SET ������ɫ='' WHERE ������ɫ=null"
Data1.Database.Execute "UPDATE JHCK SET ��������='' WHERE ��������=null"
Data1.Database.Execute "UPDATE JHCK SET ���Ϲ��='' WHERE ���Ϲ��=null"

Data1.RecordSource = "SELECT JHCK.����,JHCK.��������,JHCK.���Ϲ��,JHCK.������ɫ,SUM(JHCK.�ƻ���) AS �ƻ���,SUM(JHCK.ʵ����) AS ������,SUM(JHCK.�ƻ���-JHCK.ʵ����) AS Ƿ���� FROM JHCK  GROUP BY JHCK.����,JHCK.��������,JHCK.���Ϲ��,JHCK.������ɫ"
Data1.Refresh
Call SX2(Data1, MSFlexGrid1, 5)
Call SX2(Data1, MSFlexGrid1, 6)
Call SX2(Data1, MSFlexGrid1, 7)

Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.���ݺ�='" & Text2.Text & "' "
Data6.Refresh

End Sub


Private Sub Command9_Click()
On Error Resume Next
Text2.Text = "00000001"

Data8.RecordSource = "SELECT MAX(VAL(KPD.���ݺ�)) FROM KPD"
Data8.Refresh
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

End Sub


Private Sub DBCombo2_Click(Area As Integer)
If DBCombo2.Text = "" Then
Data1.RecordSource = "SELECT * FROM KPD WHERE KPD.�ƻ����� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND KPD.OK<>'Y'"
Data1.Refresh
Data2.RecordSource = "SELECT KPD.��ǩ FROM KPD WHERE KPD.�ƻ����� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') AND KPD.OK<>'Y' GROUP BY KPD.��ǩ"
Data2.Refresh

Else
Data1.RecordSource = "SELECT * FROM KPD WHERE KPD.�ƻ����� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') and kpd.�ͻ�='" & DBCombo2.Text & "'AND KPD.OK<>'Y'"
Data1.Refresh
Data2.RecordSource = "SELECT KPD.��ǩ FROM KPD WHERE KPD.�ƻ����� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "') and kpd.�ͻ�='" & DBCombo2.Text & "'AND KPD.OK<>'Y' GROUP BY KPD.��ǩ"
Data2.Refresh

End If

End Sub





Private Sub DTPicker3_Change()
Text4.Text = Month(DTPicker3.Value)
End Sub

Private Sub DTPicker3_CloseUp()
Text4.Text = Month(DTPicker3.Value)
End Sub

Private Sub Form_Load()
On Error Resume Next
DTPicker1.Value = Date - 15
DTPicker2.Value = Date
DBCombo2.Text = ""
DBCombo3.Text = ""
DBCombo4.Text = ""
Command5.Visible = False
Text2.Text = ""
Text3.Text = ""
DBCombo5.Text = ""
DTPicker3.Value = Date
For i = 0 To 11
Text1(i).Text = ""
Next
Option6.Value = True

Data10.DatabaseName = "d:\���ݿ�\htgl\2011\SCZYJHD.MDB"
Data10.RecordSource = "SELECT * FROM SCZY_ZDH WHERE ����='��ʼ'AND ���� BETWEEN CDATE('" & DTPicker1.Value & "') AND CDATE('" & DTPicker2.Value & "')"
Data10.Refresh

Data1.DatabaseName = "d:\���ݿ�\htgl\2011\SCZYJHD.MDB"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\htgl\2011\SCZYJHD.MDB"

Data3.DatabaseName = "d:\���ݿ�\htgl\2011\CKGL.MDB"
Data3.Refresh

Data4.DatabaseName = "d:\���ݿ�\htgl\2011\sczyjhd.mdb"
Data4.RecordSource = "select KHZL.���  from KHZL group by KHZL.���"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\htgl\2011\SCZYJHD.MDB"
Data5.RecordSource = "select ct.������  from ct group by ct.������ ORDER BY VAL(CT.������)"
Data5.Refresh


Data7.DatabaseName = "d:\���ݿ�\htgl\2011\CKGL.MDB"
Data7.RecordSource = "SELECT MAX(VAL(KPD.IP)) FROM KPD WHERE KPD.��ǩ='" & DBCombo1.Text & "'"
Data7.Refresh

Data8.DatabaseName = "d:\���ݿ�\htgl\2011\CKGL.MDB"
Data8.RecordSource = "SELECT MAX(VAL(KPD.���ݺ�)) FROM KPD"
Data8.Refresh


ProgressBar1.Visible = False
Timer1.Enabled = False
Text2.Enabled = False
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Data6.DatabaseName = "d:\���ݿ�\htgl\2011\CKGL.MDB"
Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.���ݺ�='" & Text2.Text & "' "
Data6.Refresh

MSFlexGrid4.ColWidth(0) = 400
MSFlexGrid4.ColWidth(1) = 0
MSFlexGrid4.ColWidth(2) = 0
MSFlexGrid4.ColWidth(3) = 0
MSFlexGrid4.ColWidth(4) = 0
MSFlexGrid4.ColWidth(5) = 0
MSFlexGrid4.ColWidth(6) = 0
MSFlexGrid4.ColWidth(7) = 0
MSFlexGrid4.ColWidth(8) = 1500
MSFlexGrid4.ColWidth(9) = 1200
MSFlexGrid4.ColWidth(10) = 0
MSFlexGrid4.ColWidth(11) = 0

MSFlexGrid2.ColWidth(0) = 400
MSFlexGrid2.ColWidth(1) = 1200
MSFlexGrid2.ColWidth(2) = 1200
MSFlexGrid2.ColWidth(3) = 1200

MSFlexGrid3.ColWidth(1) = 1200

For i = 10 To 17
MSFlexGrid2.ColWidth(i) = 0
Next

MSFlexGrid3.ColWidth(3) = 1200

Text4.Text = Month(Date)
Select Case Text4.Text
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


DBCombo1.Text = ""
End Sub



Private Sub Label6_Click()
Text2.Enabled = False
End Sub

Private Sub Label6_DblClick()
Text2.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
rs = MSFlexGrid2.Row
Data3.Recordset.MoveFirst
Data3.Recordset.Move rs - 1
Text3.Text = Data3.Recordset.Fields(5)
For i = 0 To 6
Text1(i).Text = Data3.Recordset.Fields(i)
Next
Text1(8).Text = Data3.Recordset.Fields(7)
Text1(5).Text = 0
Text1(9).Text = Date

Data7.RecordSource = "SELECT MAX(VAL(KPD.���)) FROM KPD WHERE KPD.��ǩ='" & DBCombo1.Text & "'"
Data7.Refresh
Text1(10).Text = 1
If Data7.Recordset.EOF Then
Text1(10).Text = 1
Else
Text1(10).Text = Data7.Recordset.Fields(0) + 1
End If


End Sub

Private Sub MSFlexGrid3_DBLClick()
On Error Resume Next
rs = MSFlexGrid3.Row
Data6.Recordset.MoveFirst
Data6.Recordset.Move rs - 1
Text1(5).Text = Data6.Recordset.Fields(8)


DBCombo2.Text = Data6.Recordset.Fields(0)
 DBCombo4.Text = Data6.Recordset.Fields(1)
DBCombo1.Text = Data6.Recordset.Fields(2)
Text1(0).Text = Data6.Recordset.Fields(3)
 Text1(1).Text = Data6.Recordset.Fields(4)
 Text1(2).Text = Data6.Recordset.Fields(5)
 Text1(3).Text = Data6.Recordset.Fields(6)
 Text1(4).Text = Data6.Recordset.Fields(7)
Text1(5).Text = Data6.Recordset.Fields(8)
 Text1(6).Text = Data6.Recordset.Fields(9)
 Text1(7).Text = Data6.Recordset.Fields(10)
 DBCombo3.Text = Data6.Recordset.Fields(11)
 Text1(9).Text = Data6.Recordset.Fields(12)
 Text1(10).Text = Data6.Recordset.Fields(13)
 Text2.Text = Data6.Recordset.Fields(14)
 Text1(11).Text = Data6.Recordset.Fields(15)
 Text1(8).Text = Data6.Recordset.Fields(16)
 DBCombo5.Text = Data6.Recordset.Fields(17)

End Sub

Private Sub MSFlexGrid4_dblClick()
On Error Resume Next
rs = MSFlexGrid4.Row
If Data10.Recordset.EOF Then Exit Sub
Data10.Recordset.MoveFirst
Data10.Recordset.Move rs - 1
DBCombo4.Text = Data10.Recordset.Fields(7)

End Sub

Private Sub Option1_Click()
KB = "�����"
Data8.Database.Execute "DELETE * FROM CLRCZZLS"
Data8.Database.Execute "DELETE * FROM CLRCZZHZLS"
Data8.Database.Execute "INSERT INTO CLRCZZLS(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ,����,KPD.����,KPD.����,KPD.���� from KPD WHERE KPD.����='" & DBCombo4.Text & "' AND KPD.���='�����'  AND KPD.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data8.Database.Execute "UPDATE CLRCZZLS SET ���='����',����=-���� where ���=NULL"
Data8.Database.Execute "INSERT INTO CLRCZZLS(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select CKBL.��������,CKBL.���Ϲ��,CKBL.���ϵ�λ,CKBL.��ɫ,����,CKBL.����,CKBL.����,CKBL.���� from ckBL WHERE CKBL.����='" & DBCombo4.Text & "' AND CKBL.���='�����'  AND CKBL.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data8.Database.Execute "UPDATE CLRCZZLS SET ���='���' WHERE ���=NULL"
Data8.Database.Execute "INSERT INTO CLRCZZHZLS(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����) SELECT CLRCZZLS.����,CLRCZZLS.��������,CLRCZZLS.���Ϲ��,CLRCZZLS.���ϵ�λ,CLRCZZLS.��ɫ,����,SUM(CLRCZZLS.����) AS L,AVG(CLRCZZLS.����) AS D FROM CLRCZZLS GROUP BY CLRCZZLS.����,CLRCZZLS.��������,CLRCZZLS.���Ϲ��,CLRCZZLS.���ϵ�λ,CLRCZZLS.��ɫ,����"
Data3.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,���� FROM CLRCZZHZLS WHERE CLRCZZHZLS.����>0 ORDER BY CLRCZZHZLS.����"
Data3.Refresh
Call SX2(Data3, MSFlexGrid2, 6)
Call SX2(Data3, MSFlexGrid2, 7)
End Sub

Private Sub MSF()
With MSFlexGrid3
    c = .Col: r = .Row    '''''C�У���R��
        Text1111.Left = .Left + .ColPos(c)
        Text1111.Top = .Top + .RowPos(r)
        Text1111.Width = .ColWidth(c)
        Text1111.Height = .RowHeight(r)
        Text1111 = .Text
        Text1111.Visible = True
        Text1111.SetFocus
End With
End Sub

Private Sub MSFlexGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub


Private Sub Option5_Click()
KB = "�ɹ����"
Data8.Database.Execute "DELETE * FROM KCCXLS"
Data8.Database.Execute "DELETE * FROM KCCXHZLS"
Data8.Database.Execute "INSERT INTO KCCXLS(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.��ɫ,CKGL.����,CKGL.����,CKGL.����,CKGL.���� from ckgl WHERE CKGL.����='" & DBCombo4.Text & "' AND CKGL.���<>'�����' AND  CKGL.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data8.Database.Execute "UPDATE KCCXLS SET ���='���' where ���=NULL"
Data8.Database.Execute "INSERT INTO KCCXLS(��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����) select KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ,KPD.����,KPD.����,KPD.����,KPD.���� from KPD WHERE KPD.����='" & DBCombo4.Text & "'  AND KPD.���� BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') "
Data8.Database.Execute "UPDATE KCCXLS SET ���='����',����=-���� WHERE ���=NULL"
Data8.Database.Execute "INSERT INTO KCCXHZLS(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����) SELECT KCCXLS.����,KCCXLS.��������,KCCXLS.���Ϲ��,KCCXLS.���ϵ�λ,KCCXLS.��ɫ,KCCXLS.����,SUM(KCCXLS.����) AS L,AVG(KCCXLS.����) AS D FROM KCCXLS GROUP BY KCCXLS.����,KCCXLS.��������,KCCXLS.���Ϲ��,KCCXLS.���ϵ�λ,KCCXLS.��ɫ,KCCXLS.����"
Data3.RecordSource = "SELECT ��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,���� FROM KCCXHZLS WHERE KCCXHZLS.����>0 ORDER BY KCCXHZLS.����"
Data3.Refresh
Call SX2(Data3, MSFlexGrid2, 7)
Call SX2(Data3, MSFlexGrid2, 6)
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
       Case 5
       Text1(7).Text = Format(Val(Text1(5).Text) * Val(Text1(6).Text), "#0.00")
End Select
End Sub

Private Sub Text1111_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then
    MSFlexGrid3.Text = Text1111.Text
    Text1111.Visible = False
    MSFlexGrid3.SetFocus
End If
End Sub

Private Sub Text1111_LostFocus()
On Error Resume Next
Data6.Recordset.MoveFirst
Data6.Recordset.Move r - 1
Data6.Recordset.Edit

Data6.Recordset.Fields(c - 1) = Text1111.Text
Data6.Recordset.Update

Text1111.Visible = False
MSFlexGrid3.SetFocus
End Sub


Private Sub Text2_Change()
Data6.RecordSource = "SELECT * FROM KPD WHERE KPD.���ݺ�='" & Text2.Text & "' "
Data6.Refresh
End Sub

Private Sub Text4_Change()
Select Case Text4.Text
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
DTPicker1.Value = K1
DTPicker2.Value = K2
End Sub

Private Sub Timer1_Timer()
If BAR = 100 Then
DataEnvironment1.Command3 Text2.Text
DataReport9.Show 1
DataEnvironment1.rsCommand3.Close
Timer1.Enabled = False
ProgressBar1.Visible = False
Text2.Text = "00000001"
If Data8.Recordset.EOF Then
Text2.Text = "00000001"
Else
Text2.Text = Left("00000000", 8 - Len(Trim(Str(Data8.Recordset.Fields(0) + 1)))) + Trim(Str(Data8.Recordset.Fields(0) + 1))
End If

Exit Sub
End If
BAR = BAR + 1
ProgressBar1.Value = BAR
End Sub

