VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy29 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����ϱ���"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form29"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Data Data17 
      Caption         =   "Data17"
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
      Width           =   4575
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӱ���"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Data Data16 
      Caption         =   "Data16"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10440
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   960
      Width           =   855
   End
   Begin VB.Data Data15 
      Caption         =   "Data15"
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
      Top             =   10560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   56
      Text            =   "Text2"
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������鿴"
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����ϲ鿴"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
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
      Top             =   10440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ɫ�鿴"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7680
      Width           =   1455
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "Formy29.frx":0000
      Height          =   330
      Left            =   960
      TabIndex        =   49
      Top             =   7680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "YS"
      Text            =   "DBCombo2"
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����Ϣ"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10560
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸�"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy29.frx":0015
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5160
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   2
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy29.frx":0029
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "���"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      DataSource      =   "Data2"
      Height          =   330
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formy29.frx":003D
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      ListField       =   "��������"
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   4
      Left            =   840
      TabIndex        =   11
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   330
      Index           =   5
      Left            =   3360
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      DataSource      =   "Data3"
      Height          =   330
      Index           =   9
      Left            =   3360
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   12648447
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   "DBCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   10680
      Visible         =   0   'False
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
      Left            =   2280
      TabIndex        =   15
      Top             =   10560
      Visible         =   0   'False
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
   Begin MSDBCtls.DBCombo DBCombo1 
      Height          =   240
      Index           =   23
      Left            =   2880
      TabIndex        =   16
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   423
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   ""
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   5.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formy29.frx":0051
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   18
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Formy29.frx":0066
      Height          =   6015
      Left            =   7800
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
      FormatString    =   "��¼�� "
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "���ϱ���"
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   7575
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text3"
         Top             =   720
         Width           =   1815
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   6
         Left            =   3240
         TabIndex        =   19
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   8
         Left            =   3240
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formy29.frx":007B
         Height          =   330
         Index           =   10
         Left            =   4920
         TabIndex        =   21
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3240
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formy29.frx":008F
         Height          =   330
         Index           =   11
         Left            =   4920
         TabIndex        =   39
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   "XM"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formy29.frx":00A3
         Height          =   330
         Index           =   12
         Left            =   4920
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "Formy29.frx":00B7
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   4920
         TabIndex        =   41
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   "MC"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   4920
         TabIndex        =   43
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Height          =   330
         Index           =   15
         Left            =   6360
         TabIndex        =   44
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "gg yy/MM/dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   6360
         TabIndex        =   45
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   12648447
         ListField       =   ""
         Text            =   "DBCombo1"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ԭ����"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   47
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   46
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   42
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         Height          =   375
         Index           =   15
         Left            =   6360
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   14
         Left            =   6360
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ע"
         Height          =   375
         Index           =   11
         Left            =   4320
         TabIndex        =   34
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   32
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   10
         Left            =   2640
         TabIndex        =   31
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ɫ"
         Height          =   375
         Index           =   9
         Left            =   2640
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��λ"
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�ͻ�"
         Height          =   375
         Index           =   0
         Left            =   6840
         TabIndex        =   23
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   33
         Top             =   1680
         Width           =   615
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   8880
      TabIndex        =   60
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   11520
      TabIndex        =   61
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   81002497
      CurrentDate     =   39177
   End
   Begin VB.Line Line1 
      X1              =   10920
      X2              =   11280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "���ڷ�Χ"
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
      Left            =   7800
      TabIndex        =   62
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "7"
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
      Index           =   18
      Left            =   3120
      TabIndex        =   59
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "7"
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
      Index           =   16
      Left            =   0
      TabIndex        =   58
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ǰ�Ĳ�������"
      Height          =   375
      Index           =   6
      Left            =   1200
      TabIndex        =   55
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      Caption         =   "��ɫ"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   50
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�� �� �� �� �� �� "
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   38
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      Caption         =   "�������"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   37
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Formy29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X, BAR, SHX, CZBZH, CZSX, SJBL As Integer ''''''''CZBZH������־
Dim ba As Database: Dim rr As Integer
Dim rs As Single: Dim rd As Recordset: Dim ba1 As Database: Public ll As Integer: Public K1, K2 As String
Dim rd1 As Recordset
Dim A As String  '�м����
Dim B As Double
Dim c As Integer
Dim kg As Integer
Dim bb As Long
Dim cc As String
Dim kkf As Integer
Dim N As Integer
Dim DH  As Integer
Dim fh As String

Private Sub Command12_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command1_Click()
Data2.Database.Execute "DELETE * FROM CLRCZZ"
Data2.Database.Execute "DELETE * FROM CLRCZZHZ"
Data2.Database.Execute "INSERT INTO CLRCZZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����,����) select ����,CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.��ɫ,CKGL.����,CKGL.����,CKGL.����,CKGL.����,��� from ckgl WHERE ���='�����' AND CKGL.���� BETWEEN CDATE('" & DTPicker3.Value & "') AND CDATE('" & DTPicker4.Value & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET ���='���' where ���=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����,����) select ����,CKGL.��������,CKGL.���Ϲ��,CKGL.���ϵ�λ,CKGL.��ɫ,CKGL.����,CKGL.����,CKGL.����,CKGL.����,��� from ckgl WHERE ���='�ɹ����' AND CKGL.���� BETWEEN CDATE('" & DTPicker3.Value & "') AND CDATE('" & DTPicker4.Value & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET ���='���' where ���=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����,����) select ��ע,CKBL.��������,CKBL.���Ϲ��,CKBL.���ϵ�λ,CKBL.��ɫ,CKBL.����,CKBL.����,CKBL.����,CKBL.����,��� from ckBL WHERE CKBL.���� BETWEEN CDATE('" & DTPicker3.Value & "') AND CDATE('" & DTPicker4.Value & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET ���='����',����=-���� WHERE ���=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����,����) select ����,KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ,KPD.����,KPD.����,KPD.����,KPD.����,��� from KPD WHERE ��ǩ<>'�����' and  KPD.���� BETWEEN CDATE('" & DTPicker3.Value & "') AND CDATE('" & DTPicker4.Value & "') "
Data2.Database.Execute "INSERT INTO CLRCZZ(����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,����,����) select ��ע,KPD.��������,KPD.���Ϲ��,KPD.���ϵ�λ,KPD.��ɫ,KPD.����,KPD.����,KPD.����,KPD.����,��� from KPD WHERE ��ǩ='�����' and KPD.���� BETWEEN CDATE('" & DTPicker3.Value & "') AND CDATE('" & DTPicker4.Value & "') "
Data2.Database.Execute "UPDATE CLRCZZ SET ���='����',����=-���� WHERE ���=NULL"
Data2.Database.Execute "INSERT INTO CLRCZZHZ(����,����,����,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����) SELECT ����,����,CLRCZZ.����,CLRCZZ.��������,CLRCZZ.���Ϲ��,CLRCZZ.���ϵ�λ,CLRCZZ.��ɫ,CLRCZZ.����,����,format(SUM(CLRCZZ.����),'#0.00') FROM CLRCZZ GROUP BY ����,����,CLRCZZ.����,CLRCZZ.��������,CLRCZZ.���Ϲ��,CLRCZZ.���ϵ�λ,CLRCZZ.��ɫ,CLRCZZ.����,����"
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����>0 order by ����"
Data13.Refresh
End Sub

Private Sub Command10_Click()
Data15.RecordSource = "SELECT DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������,format(SUM(DHCLB.��������),'#0.00') AS ������ FROM DHCLB WHERE DHCLB.����='" & Text2.Text & "' GROUP BY DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������"
Data15.Refresh
End Sub

Private Sub Command3_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.��������='" & DBCombo1(3).Text & "' AND CLRCZZHZ.��ɫ='" & DBCombo2.Text & "' AND CLRCZZHZ.����>0"
Data13.Refresh
End Sub

Private Sub Command33_Click()
Unload Me
End Sub


Private Sub Command5_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE INSTR(CLRCZZHZ.��������,'" & DBCombo1(3).Text & "')>0 AND CLRCZZHZ.����>0"
Data13.Refresh
End Sub

Private Sub Command11_Click()
On Error Resume Next
If CZBZH = 0 Then
If MsgBox("����ѡ��ֿ�����Ϣ,������?", vbYesNo) = vbNo Then Exit Sub
End If

If DBCombo1(11).Text = "" Then
MsgBox ("��ѡ�����ˣ�")
Exit Sub
End If

If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("���������۲���Ϊ�գ���")
Exit Sub
End If

If Val(DBCombo1(8).Text) > Val(Text1.Text) Then
MsgBox ("����������������������")
Exit Sub
End If


rd.AddNew
For i = 0 To rd.Fields.Count - 1
rd.Fields(i) = DBCombo1(i).Text
Next
rd.Fields(14) = Text3.Text
rd.Update

'For i = 3 To RD.Fields.Count - 3
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next

'Data13.RecordSource = "select  * from ckgl WHERE CKgl.��������='" & DBCombo1(3).Text & "' and CKGL.����>CKGL.ʵ���� and CKGL.����='" & DBCombo1(15).Text & "'  order by Val(ckgl.���)"
Data13.Refresh
Data5.RecordSource = "select * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
Data5.Refresh
Data7.RecordSource = "select MAX(VAL(CKbl.���)) from ckbl "
Data7.Refresh
If Data5.Recordset.EOF Then
MSFlexGrid1.TextMatrix(0, 0) = "��¼��"
End If
Data5.Recordset.MoveLast
MSFlexGrid1.TextMatrix(0, 0) = "��¼��"
For i = 1 To Data5.Recordset.RecordCount
MSFlexGrid1.TextMatrix(i, 0) = i
Next
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date

CZBZH = 0
End Sub

Private Sub Command2_Click()
On Error Resume Next

If DBCombo1(11).Text = "" Then
MsgBox ("��ѡ�����ˣ�")
Exit Sub
End If

If DBCombo1(8).Text = "" Or DBCombo1(9).Text = "" Then
MsgBox ("���������۲���Ϊ�գ���")
Exit Sub
End If

Data5.Recordset.Edit
For i = 0 To Data5.Recordset.Fields.Count - 1
Data5.Recordset.Fields(i) = DBCombo1(i).Text
Next
Data5.Recordset.Fields(14) = Text3.Text
Data5.Recordset.Update


'For i = 3 To RD.Fields.Count - 3
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next

Data5.RecordSource = "select   * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
Data5.Refresh

Data7.RecordSource = "select   MAX(VAL(CKGL.���)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
DBCombo1(0).SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next

Data5.Recordset.Delete

'For i = 3 To RD.Fields.Count - 2
'If i = 18 Then
'CWY = DBCombo1(i).Text
'End If
'DBCombo1(i).Text = ""
'Next


Data5.RecordSource = "select   * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
Data5.Refresh
Data7.RecordSource = "select   MAX(VAL(CKGL.���)) from ckgl "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
DBCombo1(0).SetFocus
End Sub


Private Sub Command8_Click()
On Error Resume Next
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
DBCombo1(15).Text = Date
DBCombo1(8).Text = 0
SHX = 0
Data1.Refresh
Data3.Refresh
Data4.Refresh
Data6.Refresh
Data8.Refresh
Data9.Refresh
Data7.Database.Execute "UPDATe CKBL SET ���='0'  WHERE ���=null"
Data7.RecordSource = "select   MAX(VAL(CKBL.���)) from CKBL "
Data7.Refresh
DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1
DBCombo1(15).Text = Date
Data5.RecordSource = "select   * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
Data5.Refresh
End Sub


Private Sub Command9_Click()
Data13.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����='" & DBCombo1(12).Text & "' AND CLRCZZHZ.����>0"
Data13.Refresh
End Sub

Private Sub DBCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 1
       Data5.RecordSource = "select   * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
       Data5.Refresh
       Case 8
       DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")
       Case 9
       DBCombo1(9).Text = Format(DBCombo1(9).Text, "#0.00")
End Select
End Sub

Private Sub DBCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 1
       Data5.RecordSource = "select   * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
       Data5.Refresh
       Case 8
       DBCombo1(10).Text = Format(Val(DBCombo1(8).Text) * Val(DBCombo1(9).Text), "#0.00")
       Case 9
       DBCombo1(9).Text = Format(DBCombo1(9).Text, "#0.00")
End Select

End Sub

Private Sub DBCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value
Text4.SetFocus
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_CloseUp()
Text5.Text = DTPicker2.Value
Text5.SetFocus
End Sub

Private Sub Form_Load()
Me.Caption = Me.Caption + "��Ȳ��� " + LJB
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
DTPicker3.Value = Date
DTPicker4.Value = Date
Text3.Text = ""

Set ba = OpenDatabase("d:\���ݿ�\\htgl\2011\ckgl.MDB")
Set rd = ba.OpenRecordset("ckBL", dbOpenDynaset)

For i = 0 To rd.Fields.Count - 1
DBCombo1(i).Text = ""
Next
DBCombo2.Text = ""
DBCombo1(15).Text = Date
DBCombo1(16).Enabled = False
DBCombo1(8).Text = 0
DBCombo1(13).Text = "�����"


Data1.DatabaseName = "d:\���ݿ�\\htgl\2011\sczyjhd.mdb"
Data1.RecordSource = "select ��� from KHZL group by ���"
Data1.Refresh

Data2.DatabaseName = "d:\���ݿ�\\htgl\2011\CKGL.mdb"
Data2.RecordSource = "select CKGL.�������� from ckgl WHERE CKGL.���<>'�ɹ����' group by ckgl.��������"
Data2.Refresh


Data3.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data3.RecordSource = "select CW.MC from CW group by CW.MC"
Data3.Refresh

Data4.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data4.RecordSource = "select fzr.xm  from fzr group by fzr.xm"
Data4.Refresh

Data5.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data5.RecordSource = "select  * from ckBL WHERE CKBL.����='" & DBCombo1(1).Text & "' order by Val(ckBL.���) DESC"
Data5.Refresh

Data6.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data6.RecordSource = "select KL.MC from KL group by KL.MC"
Data6.Refresh



Data7.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"

Data7.Database.Execute "UPDATE CKBL SET ���='0' WHERE ���=NULL"

Data7.RecordSource = "select   MAX(VAL(CKBL.���)) from ckBL "
Data7.Refresh

DBCombo1(16).Text = 1
DBCombo1(16).Text = Data7.Recordset.Fields(0) + 1

Data8.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data8.RecordSource = "select KB.MC from KB group by KB.MC"
Data8.Refresh


Data13.DatabaseName = "d:\���ݿ�\\htgl\2011\CKGL.MDB"

Data9.DatabaseName = "d:\���ݿ�\\htgl\2011\ckgl.mdb"
Data9.RecordSource = "select CLDW.MC from CLDW group by CLDW.MC"
Data9.Refresh

Data10.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data11.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data14.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"
Data14.RecordSource = "select YS.YS  from YS group by YS.YS"
Data14.Refresh

Data15.DatabaseName = "d:\���ݿ�\\htgl\2011\SCZYJHD.mdb"

Data16.DatabaseName = "d:\���ݿ�\\htgl\2011\CKGL.mdb"

MSFlexGrid1.ColWidth(0) = 300
MSFlexGrid1.ColWidth(1) = 0
MSFlexGrid1.ColWidth(2) = 0
MSFlexGrid1.ColWidth(3) = 0
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 1000
MSFlexGrid1.ColWidth(8) = 1200
MSFlexGrid1.ColWidth(9) = 1200
MSFlexGrid1.ColWidth(10) = 0
MSFlexGrid1.ColWidth(11) = 0
MSFlexGrid1.ColWidth(14) = 1200
MSFlexGrid1.ColWidth(15) = 1200

MSFlexGrid3.ColWidth(0) = 300
MSFlexGrid2.ColWidth(1) = 2500
MSFlexGrid2.ColWidth(2) = 1500
MSFlexGrid2.ColWidth(7) = 0
MSFlexGrid2.ColWidth(10) = 1500



DBCombo1(0).TabIndex = 0

Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub






Private Sub Label3_Click(Index As Integer)
DBCombo1(3).Enabled = False
End Sub

Private Sub Label3_dblClick(Index As Integer)
DBCombo1(3).Enabled = True
End Sub

Private Sub MSFlexGrid1_dblClick()
On Error Resume Next
rs = MSFlexGrid1.Row
Data5.Recordset.MoveFirst
Data5.Recordset.Move rs - 1
For i = 0 To Data5.Recordset.Fields.Count - 1
DBCombo1(i).Text = Data5.Recordset.Fields(i)
Next
Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
If Data13.Recordset.EOF Then
Exit Sub
End If
rs = MSFlexGrid2.Row
Data13.Recordset.MoveFirst
Data13.Recordset.Move rs - 1
Text1.Text = Data13.Recordset.Fields(5)
For i = 0 To Data13.Recordset.Fields.Count - 3
DBCombo1(3 + i).Text = Data13.Recordset.Fields(i)
Next
DBCombo1(13).Text = Data13.Recordset.Fields(8)
Text3.Text = Data13.Recordset.Fields(9)
CZBZH = 1
End Sub


Private Sub MSFlexGrid3_Click()
On Error Resume Next
rs = MSFlexGrid3.Row
'If Data2.Recordset.EOF Then Exit Sub
Data15.Recordset.MoveFirst
Data15.Recordset.Move rs - 1
DBCombo1(12).Text = Data15.Recordset.Fields(0)
DBCombo1(12).Text = Data15.Recordset.Fields(0)
DBCombo1(3).Text = Data15.Recordset.Fields(1)
Data16.RecordSource = "SELECT * FROM CLRCZZHZ WHERE CLRCZZHZ.����='" & Data15.Recordset.Fields(0) & "' AND CLRCZZHZ.��������='" & Data15.Recordset.Fields(1) & "' AND ��ɫ='" & Data15.Recordset.Fields(4) & "' AND CLRCZZHZ.����>0"
Data16.Refresh
End Sub

Private Sub MSFlexGrid4_Click()
On Error Resume Next
rs = MSFlexGrid4.Row
Data16.Recordset.MoveFirst
Data16.Recordset.Move rs - 1
DBCombo1(12).Text = Data16.Recordset.Fields(7)
DBCombo1(3).Text = Data16.Recordset.Fields(0)
DBCombo2.Text = Data16.Recordset.Fields(3)
DBCombo1(1).Text = Text2.Text
End Sub

Private Sub Text2_Change()
Data15.RecordSource = "SELECT DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������,format(SUM(DHCLB.��������),'#0.00') AS ������ FROM DHCLB WHERE DHCLB.����='" & Text2.Text & "' GROUP BY DHCLB.���Ͽ���,DHCLB.��������,DHCLB.���Ϲ��,DHCLB.���ϵ�λ,DHCLB.������ɫ,DHCLB.��������"
Data15.Refresh
End Sub
