VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Forms51 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�����뵥����"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   Icon            =   "Forms51.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   375
      Left            =   8640
      Top             =   8160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   9840
      TabIndex        =   70
      Text            =   "Text7"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   15720
      TabIndex        =   68
      Text            =   "Text6"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "�뵥��ӡѡ��"
      Height          =   1575
      Left            =   10080
      TabIndex        =   64
      Top             =   4680
      Width           =   1450
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFF80&
         Caption         =   "����"
         Height          =   370
         Left            =   240
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFF80&
         Caption         =   "����"
         Height          =   370
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text5 
      Height          =   370
      Left            =   9000
      TabIndex        =   63
      Text            =   "Text5"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "¼��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1920
      TabIndex        =   52
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3960
      TabIndex        =   51
      Text            =   "Text2"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Caption         =   "�Զ�"
      Height          =   375
      Left            =   12000
      TabIndex        =   50
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9000
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   5880
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   120
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Caption         =   "�ֶ�"
      Height          =   375
      Left            =   12000
      TabIndex        =   47
      Top             =   4680
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Caption         =   "����"
      Height          =   375
      Left            =   12000
      TabIndex        =   46
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox TxtReceive 
      Height          =   375
      Left            =   10560
      MultiLine       =   -1  'True
      TabIndex        =   45
      Text            =   "Forms51.frx":440A
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtSend 
      Height          =   375
      Left            =   10560
      MultiLine       =   -1  'True
      TabIndex        =   44
      Text            =   "Forms51.frx":4411
      Top             =   8400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6960
      Top             =   120
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ɨ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "У��ë��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ǩ��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "¼���ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "72�뵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   6480
      Top             =   120
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "������ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�����Զ�"
      Enabled         =   0   'False
      Height          =   615
      Left            =   15360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "132�뵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   9840
      Top             =   10320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   10200
      Top             =   10200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   9960
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   9360
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   10200
      Top             =   10320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   12360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7560
      Top             =   9120
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10320
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forms51.frx":4418
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   17415
      _cx             =   30718
      _cy             =   3413
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forms51.frx":442D
      Height          =   4095
      Left            =   840
      TabIndex        =   1
      Top             =   5400
      Width           =   7575
      _cx             =   13361
      _cy             =   7223
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   11280
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      Format          =   328269825
      CurrentDate     =   39177
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   1
      Left            =   840
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   2
      Left            =   7800
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   3
      Left            =   5040
      TabIndex        =   13
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms51.frx":4442
      Height          =   330
      Index           =   4
      Left            =   9960
      TabIndex        =   14
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   ""
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   5
      Left            =   15480
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   6
      Left            =   3120
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   7
      Left            =   5040
      TabIndex        =   17
      Top             =   3960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   8
      Left            =   840
      TabIndex        =   18
      Top             =   3960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   9
      Left            =   17040
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   10
      Left            =   14160
      TabIndex        =   23
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms51.frx":4457
      Height          =   330
      Index           =   11
      Left            =   13800
      TabIndex        =   24
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "mc"
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   12
      Left            =   10080
      TabIndex        =   25
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "Forms51.frx":446C
      Height          =   330
      Index           =   13
      Left            =   11520
      TabIndex        =   43
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "����"
      Text            =   "DataCombo1"
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   4800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forms51.frx":4481
      Height          =   615
      Left            =   840
      TabIndex        =   54
      Top             =   9480
      Width           =   7575
      _cx             =   13361
      _cy             =   1085
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   14
      Left            =   17040
      TabIndex        =   56
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   15
      Left            =   840
      TabIndex        =   60
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   330
      Index           =   16
      Left            =   3120
      TabIndex        =   61
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ӡ����"
      Height          =   495
      Left            =   9120
      TabIndex        =   69
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "����Ա"
      Height          =   255
      Index           =   2
      Left            =   15720
      TabIndex        =   67
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����ϵ��"
      Height          =   375
      Index           =   15
      Left            =   9000
      TabIndex        =   62
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   59
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�׺�"
      Height          =   255
      Index           =   13
      Left            =   840
      TabIndex        =   58
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "��������"
      Height          =   255
      Left            =   17040
      TabIndex        =   55
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�س�"
      Height          =   375
      Index           =   12
      Left            =   9000
      TabIndex        =   49
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ע"
      Height          =   255
      Index           =   11
      Left            =   11520
      TabIndex        =   42
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   255
      Index           =   7
      Left            =   7800
      TabIndex        =   39
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   255
      Index           =   6
      Left            =   14160
      TabIndex        =   38
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ƥ��"
      Height          =   255
      Index           =   0
      Left            =   10080
      TabIndex        =   37
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   11280
      TabIndex        =   36
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ͻ�����"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   35
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ɫ����ɫ��"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   34
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���������"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   33
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ƥ��"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   32
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���߷���"
      Height          =   255
      Index           =   2
      Left            =   15480
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���"
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����"
      Height          =   255
      Index           =   9
      Left            =   5040
      TabIndex        =   29
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "����ɨ��"
      Height          =   495
      Index           =   10
      Left            =   840
      TabIndex        =   28
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "���"
      Height          =   255
      Index           =   0
      Left            =   13800
      TabIndex        =   27
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "��������"
      Height          =   255
      Left            =   17040
      TabIndex        =   26
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "Forms51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gk As Integer
    Dim SendCount  As Long     '�����ѷ����ַ���Ӧ�ֽ���
    Dim ReceiveCount  As Long  '�����ѽ����ַ���Ӧ�ֽ���
    Dim PortSwitch As Boolean    '���崮���Ƿ�򿪱�־
    Public L As String
    Dim DisplayFlag As Boolean   '������մ����Ƿ������ʾ��־
Dim ZHT  As String
Dim xh As Integer      ''''''ѭ��
Dim sl As Integer      ''''�Զ���ӡ
Dim dybl As Integer    '''�жϼӹ������۵ĺ�ͬ����
Dim dzcbl As Integer  '''''���ӳƱ���
Dim sssj, sswd As Single    '''ʵʱ��������
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim fs As Integer
Dim cdbhf As Integer


Private Sub Command1_Click() ''''''¼���ӡ
On Error Resume Next

If Val(DataCombo4(9)) < 0.1 Then
DataCombo4(9).SetFocus
Exit Sub
End If

If DataCombo4(9).Text = "0" Then
MsgBox ("��ѡ���Σ�")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("��������ţ�")
Exit Sub
End If

DataCombo4(6) = Int(Val(DataCombo4(6)))

If Val(DataCombo4(7)) = 0 Then
MsgBox ("��������ȷ��ë��ƥ��������")
Exit Sub
End If

Timer2.Enabled = False

If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If


Adodc8.RecordSource = "select * from bmd where ����='" & DataCombo4(1) & "' and �׺�='" & DataCombo4(15) & "' and ƥ��='" & DataCombo4(12) & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
MsgBox ("���д�ƥ�ţ���ע��ƥ�ű��")
Call Command6_Click
If Option1.value = False Then
Timer2.Enabled = True
'Timer3.Enabled = True
sl = 1
dzcbl = 1
End If
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "mpbmdlr('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','����','" & DataCombo4(13).Text & "','" & DataCombo4(14).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "'  order by ƥ�� desc"
Adodc1.Refresh
'fs = Val(Text7.Text)    '''''��ӡ����
'If fs <= 0 Then
       ' MsgBox "��ӡ�����������0"
       ' Exit Sub
   ' End If
   ' For i = 1 To fs
Call dbq(Adodc5, DataCombo4(1).Text, DataCombo4(12).Text, DataCombo4(15).Text, Text7.Text)
 'Next i

Adodc2.RecordSource = "select max(ƥ��) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "' "
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc7.RecordSource = "select count(ƥ��) as �ϼ�ƥ��,round(sum(��������),2) as �ϼ����� from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "' "
Adodc7.Refresh
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
Timer2.Enabled = True
End If
DataCombo4(13) = ""
DataCombo4(9).SetFocus
End Sub

Private Sub Command10_Click()

If Option4.value = True Then
Call dmd100(Adodc5, DataCombo4(1), DataCombo4(15))
End If

If Option5.value = True Then
Call dmd100ms(Adodc5, DataCombo4(1), DataCombo4(15))
End If

Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "' and ����='����' order by ƥ�� desc"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
'Set g_Cmd = New Command
'    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
'    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
'    g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' ��ʾ�����ĸ��洢����
 '   g_Cmd.Execute           ' ִ�д洢����
'    g_Cmd.Cancel
End If
End Sub

Private Sub Command11_Click()
sql1 = "update bmd_mdxz set �뵥����=ë������,�뵥ƥ��=ë��ƥ�� where ����='" & DataCombo4(1).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
MsgBox ("У���ɹ���")
Adodc1.Refresh
End Sub

Private Sub Command12_Click()
On Error Resume Next
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Forms511.Text12 = bzgrbh
Forms511.Text2 = DataCombo4(1) + "J"
Forms511.Show
End Sub

Private Sub Command13_Click()
On Error Resume Next
Timer2.Enabled = False

If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If

If DataCombo4(9).Text = "0" Then
MsgBox ("��ѡ���Σ�")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("��������ţ�")
Exit Sub
End If

If DataCombo4(4).Text = "" Then
MsgBox ("�������������")
Exit Sub
End If



Adodc8.RecordSource = "select * from bmd where ����='" & DataCombo4(1) & "' and �׺�='" & DataCombo4(15) & "' and ƥ��='" & DataCombo4(12) & "'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
MsgBox ("���д�ƥ�ţ���ע��ƥ�ű��")
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "mpbmdlr('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','����','" & DataCombo4(13).Text & "','" & DataCombo4(14).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"      ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
    
Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "'  order by ƥ�� desc"
Adodc1.Refresh

Adodc2.RecordSource = "select max(ƥ��) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "' "
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc7.RecordSource = "select count(ƥ��) as �ϼ�ƥ��,round(sum(��������),2) as �ϼ����� from bmd where  ����='" & DataCombo4(1).Text & "' and Ʒ��='" & DataCombo4(3).Text & "' and ���߷���='" & DataCombo4(5) & "' and ����='" & DataCombo4(10) & "'"
Adodc7.Refresh

DataCombo4(9).SetFocus
Timer2.Enabled = True

End Sub

Private Sub Command2_Click()
On Error Resume Next
If DataCombo4(9).Text = "0" Then
MsgBox ("��ѡ���Σ�")
Exit Sub
End If

If DataCombo4(1).Text = "" Then
MsgBox ("��������ţ�")
Exit Sub
End If

DataCombo4(6) = Int(Val(DataCombo4(6)))

If Val(DataCombo4(7)) = 0 Then
MsgBox ("��������ȷ��ë��ƥ��������")
Exit Sub
End If

Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "mpbmdxg('" & DataCombo4(0).Text & "','" & DataCombo4(1).Text & "','" & DataCombo4(2).Text & "','" & DataCombo4(3).Text & "','" & DataCombo4(4).Text & "','" & DataCombo4(5).Text & "','" & DataCombo4(6).Text & "','" & DataCombo4(7).Text & "','" & DataCombo4(8).Text & "','" & DataCombo4(9).Text & "','" & DataCombo4(10).Text & "','" & DataCombo4(11).Text & "','" & DataCombo4(12).Text & "','" & DTPicker4.value & "','����','" & DataCombo4(13).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16).Text & "','" & Text6 & "')"      ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

Call Command6_Click

DataCombo4(13) = ""
DataCombo4(9).SetFocus
End Sub


Private Sub Command3_Click()
On Error Resume Next
If MsgBox("ȷ��ɾ����", vbYesNo) = vbNo Then Exit Sub
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "mpbmdsc1('" & DataCombo4(1).Text & "','" & DataCombo4(12).Text & "','" & DataCombo4(15).Text & "','" & DataCombo4(16) & "')"    ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel

Call Command6_Click
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
DataCombo4(9).SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click() '''��ǩ��ӡ
On Error Resume Next
'fs = Val(Text7.Text)    '''''��ӡ����
'If fs <= 0 Then
        'MsgBox "��ӡ�����������0"
       ' Exit Sub
    'End If
    'For i = 1 To fs
Call dbq(Adodc5, DataCombo4(1).Text, DataCombo4(12).Text, DataCombo4(15).Text, Text7.Text)
 'Next i
Adodc2.RecordSource = "select max(ƥ��) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "'"
Adodc2.Refresh

DataCombo4(9).Text = "0"
DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If

Adodc1.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

DataCombo4(9).SetFocus
End Sub

Private Sub Command6_Click()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "'  order by ƥ�� desc"
Adodc1.Refresh

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select max(ƥ��) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

If Option1.value = True Then
DataCombo4(9).Text = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If
DataCombo4(9).SetFocus
Adodc7.RecordSource = "select count(ƥ��) as �ϼ�ƥ��,round(sum(��������),2) as �ϼ����� from bmd where  ����='" & DataCombo4(1).Text & "' and Ʒ��='" & DataCombo4(3).Text & "' and ���߷���='" & DataCombo4(5) & "' and ����='" & DataCombo4(10) & "'"
Adodc7.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command7_Click()
If Option4.value = True Then
Call dmd(Adodc5, Adodc9, DataCombo4(1), DataCombo4(15))
End If

If Option5.value = True Then
Call dmdms(Adodc5, DataCombo4(1), DataCombo4(15))
End If

Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "' and ����='����'  order by ƥ�� desc"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
'Set g_Cmd = New Command
 '   g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
 '   g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
 '   g_Cmd.CommandText = "MPbmzk('" & DataCombo4(1).Text & "','" & DataCombo4(15).Text & "')"   ' ��ʾ�����ĸ��洢����
 '   g_Cmd.Execute           ' ִ�д洢����
 '   g_Cmd.Cancel
End If
End Sub

Private Sub Command8_Click()
Forms509.Text1(4) = DataCombo4(11)
Forms509.Check2(4).value = 1
Forms509.Check2(0).value = 1
Forms509.Timer1.Enabled = True
Forms509.Show
End Sub

Private Sub Command9_Click()
Timer1.Enabled = True
End Sub

Private Sub DataCombo4_Change(Index As Integer)
On Error Resume Next
Select Case Index
Case 1
If InStr(DataCombo4(1).Text, "J") > 0 Or InStr(DataCombo4(1).Text, "j") > 0 Then

DataCombo4(1).Text = Mid(DataCombo4(1), 1, Len(DataCombo4(1).Text) - 1)
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select �ͻ�����,����,���,Ʒ��,ͼ��,���߷���,isnull(���ƥ��,0) as ƥ��,isnull(�������,0) as ����,ɫ��+ɫ�� as ɫ��,����,���,'' as �ܱ�ע,'' as ��ͬ����,�׺�,��� from v_kpd_ok where ����='" & DataCombo4(1).Text & "' and ���<>'ԲͲ'  ORDER BY ���"
Adodc3.Refresh

VSFlexGrid2.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid2.AutoSize 0, VSFlexGrid2.Cols - 1, False, 30

If Adodc3.Recordset.EOF Then
For i = 0 To 10
If i = 1 Then i = i + 1
DataCombo4(i).Text = ""
Next
Else
For i = 2 To 8
DataCombo4(i).Text = Adodc3.Recordset.Fields(i)
Next
If DataCombo4(6) <> Int(DataCombo4(6)) Then
DataCombo4(6) = Int(Val(DataCombo4(6))) + 1
End If
''DataCombo4(5) = Val(DataCombo4(5)) * 100
DataCombo4(0).Text = Adodc3.Recordset.Fields(0)
DataCombo4(4).Text = Adodc3.Recordset.Fields(10)
DataCombo4(10).Text = Adodc3.Recordset.Fields(9)
DataCombo4(15).Text = Adodc3.Recordset.Fields(13)
DataCombo4(16).Text = Adodc3.Recordset.Fields(14)
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(9).SetFocus
End If

Else
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select �ͻ�����,����,���,Ʒ��,ͼ��,���߷���,isnull(���ƥ��,0) as ƥ��,isnull(�������,0) as ����,ɫ��+ɫ�� as ɫ��,����,���,'' as �ܱ�ע,'' as ��ͬ����,�׺�,��� from v_kpd_ok where ����='" & DataCombo4(1).Text & "' and ���<>'ԲͲ'  ORDER BY ���"
Adodc3.Refresh

VSFlexGrid2.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid2.AutoSize 0, VSFlexGrid2.Cols - 1, False, 30

End If

Case 16

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "'  order by ƥ�� desc"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select isnull(max(ƥ��),0) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If


Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select count(ƥ��) as �ϼ�ƥ��,round(sum(��������),2) as �ϼ����� from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "' "
Adodc7.Refresh

Case 9
If Option3.value = True And Val(DataCombo4(9)) = Val(Text1) Then
sssj = 0
xh = 1
Timer4.Enabled = True
End If

DataCombo4(14) = Format(Val(DataCombo4(9)) * Val(Text5), "#0.00")

End Select

End Sub


Private Sub DataCombo4_GotFocus(Index As Integer)
Select Case Index
       Case 1
       Case 9
DataCombo4(9) = "0.1"
End Select
End Sub

Private Sub DataCombo4_LostFocus(Index As Integer)
Select Case Index
       Case 9
If Val(DataCombo4(9)) = 0 Then
'DataCombo4(9).SetFocus
End If
End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
sql2 = "insert into yhcd(�û�,�˵�,���) values('" & yhm & "','" & Me.Caption & "','" & cdbhf & "')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.WindowState = 2
Formm1.Adodc1.Refresh
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
sql2 = "delete from yhcd where �û�='" & yhm & "' and ���='" & cdbhf & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Formm1.Adodc1.Refresh
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
       Case 10
DataCombo4(1) = ""
For i = 0 To 10
If i = 1 Then i = i + 1
DataCombo4(i).Text = ""
Next
DataCombo4(15).Text = ""
DataCombo4(16).Text = ""
Call Command6_Click
DataCombo4(1).SetFocus
End Select
End Sub

Private Sub Option1_Click()
DataCombo4(9) = "0.1"
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Option2_Click()
DataCombo4(9) = "0"
DataCombo4(9).SetFocus
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Option3_Click()
DataCombo4(9) = "0"
DataCombo4(9).SetFocus
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Text5_Change()
DataCombo4(14) = Format(Val(DataCombo4(9)) * Val(Text5), "#0.00")
End Sub

Private Sub Timer2_Timer()
On Error Resume Next                           ''''''''''''����ʹ��Ч������
If Option1.value = True Then
If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
If Err.Number = 8002 Then Exit Sub              ''''''''''''''''''''����û�ж˿ھ��˳�
        End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''���Զ�
'If jmg = "10E7661011AE6DCF" Or jmg = "10E7665011AE6E0F" Or jmg = "10E7662E11AE6DED" Or jmg = "10E7660411AE6DC3" Then   '''''''''''''''''''''''''''���ܹ�
If Option2.value = True Then
If MSComm.PortOpen = False Then
            MSComm.PortOpen = True
If Err.Number = 8002 Then Exit Sub
        End If
MSComm.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until MSComm.InBufferCount >= 16
a = MSComm.Input
TxtReceive = a
'If Mid(a, InStr(a, "="), 7) = 0 Then
b = Mid(a, InStr(a, "=") + 1, 7)
c = ""
For i = 1 To 7
c = Trim(c) + Trim(Mid(b, 8 - i, 1))
Next
TxtSend = c
clsz = Format(Val(c) + Val(Text1), "#0.0")
DataCombo4(9) = clsz
End If
'End If
''''''''''''''''''''''''''''''''''''''''''''''''ȫ�Զ�
If Option3.value = True Then
If dzcbl = 0 Then DataCombo4(9) = 0
If MSComm.PortOpen = False Then
            MSComm.PortOpen = True
If Err.Number = 8002 Then Exit Sub
        End If
MSComm.Output = Chr$(27) + "p"
 Do
       ' DoEvents
       Loop Until MSComm.InBufferCount >= 16
a = MSComm.Input
b = Mid(a, InStr(a, "=") + 1, 7)
c = ""
For i = 1 To 7
c = Trim(c) + Trim(Mid(b, 8 - i, 1))
Next
clsz = Format(Val(c) + Val(Text1), "#0.0")
'If Mid(a, 1, 2) = "=" Then
'clsz = Format(Val(Mid(a, 2, 7)) + Val(Text1), "#0.0")
DataCombo4(9) = clsz
'End If
End If
'End If                  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ܹ�
End Sub

Private Sub Timer3_Timer()
If Option3.value = True Then
If Val(DataCombo4(9)) > 1 Then
sl = sl + 1
Else
sl = 1
End If
If sl = 3 And dzcbl = 1 Then
Timer3.Enabled = False
Call Command1_Click
dzcbl = 0
xh = 1
End If
End If
End Sub

Private Sub Timer4_Timer()
If xh / 2 = Int(xh / 2) And Val(DataCombo4(9)) > 1 Then
sssj = Val(DataCombo4(9))
End If
If sssj = DataCombo4(9) Then
sswd = sswd + 1
Else
sswd = 0
End If
If sswd = 2 Then
Call Command1_Click
Timer4.Enabled = False
End If
xh = xh + 1
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Option3.value = True Or Option2.value = True Then
Exit Sub
End If

If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc1.Recordset.Move rs - 1
For i = 2 To 10
DataCombo4(i).Text = Adodc1.Recordset.Fields(i)
Next
DTPicker4.value = Adodc1.Recordset.Fields(13)
DataCombo4(12).Text = Adodc1.Recordset.Fields(12)
DataCombo4(11).Text = Adodc1.Recordset.Fields(11)
DataCombo4(0).Text = Adodc1.Recordset.Fields(0)
Text6 = Adodc1.Recordset.Fields("����")
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
rs = VSFlexGrid2.Row
Adodc3.Recordset.Move rs - 1
For i = 2 To 8
DataCombo4(i).Text = Adodc3.Recordset.Fields(i)
Next
If DataCombo4(6) <> Int(DataCombo4(6)) Then
DataCombo4(6) = Int(Val(DataCombo4(6))) + 1
End If
'''DataCombo4(5) = Val(DataCombo4(5)) * 100
DataCombo4(0).Text = Adodc3.Recordset.Fields(0)
DataCombo4(4).Text = Adodc3.Recordset.Fields(10)
DataCombo4(10).Text = Adodc3.Recordset.Fields(9)
DataCombo4(15).Text = Adodc3.Recordset.Fields(13)
DataCombo4(16).Text = Adodc3.Recordset.Fields(14)
If Option1.value = True Then
DataCombo4(9) = "0.1"
Else
DataCombo4(9) = 0
End If
DataCombo4(9).SetFocus
End Sub

Private Sub dataCombo4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub


Private Sub Form_Load()
On Error Resume Next

       TxtReceive.Text = ""
       TxtSend = ""
       MSComm.CommPort = 1
       MSComm.Settings = "9600,n,8,1"
       MSComm.InBufferSize = 1024            ' ���ý��ջ�����Ϊ1024�ֽ�
       MSComm.OutBufferSize = 1024           ' ���÷��ͻ�����Ϊ4096�ֽ�
       MSComm.InBufferCount = 0              ' ������뻺����
       MSComm.OutBufferCount = 0             ' ������������
       MSComm.SThreshold = 1                 ' ���ͻ������մ��������¼�
       MSComm.RThreshold = 1                 ' ÿX���ַ������ջ��������𴥷������¼�
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim ID1 As Long
Dim ID2 As Long
If GetID(ID1, ID2, DevicePath) = 0 Then
jmg = Hex(ID1) + Hex(ID2)           '''''''''''''''���ܹ�
Else
jmg = ""
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
Option4.value = True
Option1.value = True
Timer2.Enabled = True
dzcbl = 0
For i = 0 To 16
DataCombo4(i).Text = ""
Next
DTPicker4.value = Date
Text1.Text = "0"
Text7.Text = "1"
DataCombo4(13).Text = ""
cdbhf = cdbh
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Text5 = 0
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select * from bmd where  ����='" & DataCombo4(1).Text & "' and �׺�='" & DataCombo4(15).Text & "' order by ƥ�� desc"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select max(ƥ��) from bmd where ����='" & DataCombo4(1).Text & "' and  �׺�='" & DataCombo4(15).Text & "' "
Adodc2.Refresh

DataCombo4(12).Text = 1
If Adodc2.Recordset.EOF Then
DataCombo4(12).Text = 1
Else
DataCombo4(12).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text6 = ""
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select distinct mc,xh from bc order by xh"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "SELECT distinct ����  FROM bmdzjyy"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select * from ckgl"
Adodc9.Refresh
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

dybl = 2
Timer1.Enabled = False
VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(12) = 2000
VSFlexGrid2.ColWidth(13) = 1500

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 0
VSFlexGrid1.ColWidth(7) = 0
VSFlexGrid1.ColWidth(8) = 0
VSFlexGrid1.ColWidth(12) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(15) = 0
'''VSFlexGrid1.ColWidth(17) = 0  '''����
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 0
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(21) = 0
VSFlexGrid1.ColWidth(22) = 0

End Sub


Private Sub Timer1_Timer()
DTPicker4.value = Date
End Sub



