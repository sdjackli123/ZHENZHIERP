VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formd11111 
   BackColor       =   &H00C0E0FF&
   Caption         =   "大货配料单"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   114
      Text            =   "Combo1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "模板确认"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   3840
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   5520
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Left            =   4680
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Left            =   5160
      Top             =   10080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Left            =   5520
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Height          =   330
      Left            =   5040
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Left            =   4560
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Enabled         =   0
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
         Name            =   "宋体"
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
      Left            =   4800
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   8160
      TabIndex        =   56
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   55
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   54
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   8160
      TabIndex        =   53
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   8160
      TabIndex        =   52
      Text            =   "Text4"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   51
      Text            =   "Text4"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   50
      Text            =   "Text4"
      Top             =   10080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3840
      Width           =   855
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3840
      Width           =   855
   End
   Begin VB.Data Data33 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "客户信息"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   57
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   58
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   59
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   3
         Left            =   2040
         TabIndex        =   84
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   7
         Left            =   2520
         TabIndex        =   85
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   8
         Left            =   3000
         TabIndex        =   86
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   9
         Left            =   2040
         TabIndex        =   87
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   10
         Left            =   2520
         TabIndex        =   88
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   14
         Left            =   3000
         TabIndex        =   89
         Top             =   2640
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   15
         Left            =   2040
         TabIndex        =   90
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   19
         Left            =   2520
         TabIndex        =   91
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   20
         Left            =   3000
         TabIndex        =   92
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   21
         Left            =   2040
         TabIndex        =   93
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   22
         Left            =   2520
         TabIndex        =   94
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   23
         Left            =   3000
         TabIndex        =   95
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   24
         Left            =   600
         TabIndex        =   96
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   25
         Left            =   1080
         TabIndex        =   97
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   26
         Left            =   1560
         TabIndex        =   98
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   27
         Left            =   600
         TabIndex        =   99
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   28
         Left            =   1080
         TabIndex        =   100
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   29
         Left            =   1560
         TabIndex        =   101
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "锅号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "重量"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "料单编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确认"
      Height          =   615
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "下一编号"
      Height          =   615
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   615
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查看"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   5040
      Top             =   10080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   4800
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   5280
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   5760
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc12 
      Height          =   330
      Left            =   5400
      Top             =   10080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   330
      Left            =   4920
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   5760
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "配方单"
      Height          =   3495
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   12135
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   6
         Left            =   3480
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   5
         Left            =   3480
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   4
         Left            =   3480
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   3
         Left            =   3480
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   2
         Left            =   3480
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   1
         Left            =   3480
         TabIndex        =   105
         Text            =   "Text5"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   0
         Left            =   3480
         TabIndex        =   104
         Text            =   "Text5"
         Top             =   720
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":0000
         Height          =   330
         Index           =   0
         Left            =   6840
         TabIndex        =   73
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":0015
         Height          =   330
         Index           =   0
         Left            =   4320
         TabIndex        =   66
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "Formd11111.frx":002A
         Height          =   330
         Left            =   1200
         TabIndex        =   64
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺编号"
         Text            =   "DataCombo2"
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   0
         Left            =   7560
         TabIndex        =   47
         Text            =   "Text3"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   1
         Left            =   7560
         TabIndex        =   46
         Text            =   "Text3"
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   2
         Left            =   7560
         TabIndex        =   45
         Text            =   "Text3"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   3
         Left            =   7560
         TabIndex        =   44
         Text            =   "Text3"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   4
         Left            =   7560
         TabIndex        =   43
         Text            =   "Text3"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   5
         Left            =   7560
         TabIndex        =   42
         Text            =   "Text3"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   6
         Left            =   7560
         TabIndex        =   41
         Text            =   "Text3"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   9480
         TabIndex        =   37
         Text            =   "Text2"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   9480
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   9480
         TabIndex        =   35
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   9480
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   9480
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   1500
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   9480
         TabIndex        =   32
         Text            =   "Text2"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   9480
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   8160
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   8160
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   8160
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   8160
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1880
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   8160
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   8160
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   8160
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   720
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formd11111.frx":0040
         Height          =   330
         Index           =   12
         Left            =   1200
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺编号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formd11111.frx":0055
         Height          =   330
         Index           =   4
         Left            =   1200
         TabIndex        =   61
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺工序"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formd11111.frx":006A
         Height          =   330
         Index           =   6
         Left            =   1200
         TabIndex        =   62
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "染化助库名"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formd11111.frx":007F
         Height          =   330
         Index           =   13
         Left            =   1200
         TabIndex        =   63
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "标志"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   5
         Left            =   1200
         TabIndex        =   65
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":0095
         Height          =   330
         Index           =   1
         Left            =   4320
         TabIndex        =   67
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":00AA
         Height          =   330
         Index           =   2
         Left            =   4320
         TabIndex        =   68
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":00BF
         Height          =   330
         Index           =   3
         Left            =   4320
         TabIndex        =   69
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":00D4
         Height          =   330
         Index           =   4
         Left            =   4320
         TabIndex        =   70
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":00E9
         Height          =   330
         Index           =   5
         Left            =   4320
         TabIndex        =   71
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formd11111.frx":00FE
         Height          =   330
         Index           =   6
         Left            =   4320
         TabIndex        =   72
         Top             =   3000
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":0113
         Height          =   330
         Index           =   1
         Left            =   6840
         TabIndex        =   74
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":0128
         Height          =   330
         Index           =   2
         Left            =   6840
         TabIndex        =   75
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":013D
         Height          =   330
         Index           =   3
         Left            =   6840
         TabIndex        =   76
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":0152
         Height          =   330
         Index           =   4
         Left            =   6840
         TabIndex        =   77
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":0167
         Height          =   330
         Index           =   5
         Left            =   6840
         TabIndex        =   78
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formd11111.frx":017C
         Height          =   330
         Index           =   6
         Left            =   6840
         TabIndex        =   79
         Top             =   3000
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   11
         Left            =   10320
         TabIndex        =   80
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   16
         Left            =   10320
         TabIndex        =   81
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   17
         Left            =   10320
         TabIndex        =   82
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   18
         Left            =   10320
         TabIndex        =   83
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Bindings        =   "Formd11111.frx":0191
         Height          =   330
         Left            =   1200
         TabIndex        =   111
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "模板编号"
         Text            =   "DataCombo7"
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "模板编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   120
         TabIndex        =   112
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3480
         TabIndex        =   103
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "批次"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   7560
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "常规工艺号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   120
         TabIndex        =   40
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "压力"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   10320
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "车速"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   10320
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "次序号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   10320
         TabIndex        =   21
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工序名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "浴比"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "染化助名称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "单位"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6840
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8160
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "工艺日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   10320
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "配方编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "校值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   9480
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助代码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "染化助库"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd11111.frx":01A6
      Height          =   4575
      Left            =   240
      TabIndex        =   102
      Top             =   4680
      Width           =   16095
      _cx             =   28390
      _cy             =   8070
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
End
Attribute VB_Name = "Formd11111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BA As Database: Dim RD As Recordset
Dim c, r As Integer
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub
Private Sub Command1_Click()
Unload Me
Form221.Show
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command2_Click()

On Error Resume Next
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(10).Enabled = False
DataCombo1(11).Enabled = False
DataCombo1(12).Enabled = False

DataCombo1(13).Text = ""        '''''''''''''代码清离

If DataCombo1(0).Text = "" Or DataCombo1(1).Text = "" Or DataCombo1(2).Text = "" Then
MsgBox ("锅号，重量，料单编号不能为空")
Exit Sub
End If

For i = 0 To 6     '''''''''''''''''''''''''
If Text1(i).Text <> "" Then
DataCombo1(7).Text = DataCombo2(i).Text
DataCombo1(8).Text = DataCombo3(i).Text
DataCombo1(9).Text = Text1(i).Text
DataCombo1(10).Text = Text2(i).Text
DataCombo1(17).Text = Text4(i).Text
DataCombo1(19).Text = Text3(i).Text
Data6.Recordset.AddNew

For p = 0 To 2
Data6.Recordset.Fields(p) = DataCombo1(p).Text
Next
For p = 3 To 9
Data6.Recordset.Fields(p) = DataCombo1(p + 1).Text
Next
Data6.Recordset.Fields(18) = DataCombo1(17).Text
Data6.Recordset.Fields(15) = Data7.Recordset.RecordCount + 1
Data6.Recordset.Fields(14) = Date
Data6.Recordset.Update
Data7.Refresh
End If
Next
                '''''''''''''''''''''''
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text5(i).Text = ""
Next
DataCombo1(16).Enabled = False
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
DataCombo1(4).SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
Data7.Recordset.Edit
Data7.Recordset.Fields(3) = DataCombo1(4).Text
Data7.Recordset.Fields(4) = DataCombo1(5).Text
Data7.Recordset.Fields(5) = DataCombo1(6).Text
Data7.Recordset.Fields(6) = DataCombo2(0).Text
Data7.Recordset.Fields(7) = DataCombo3(0).Text
Data7.Recordset.Fields(8) = Text1(0).Text
Data7.Recordset.Fields(9) = Text2(0).Text
Data7.Recordset.Fields(17) = DataCombo1(16).Text
Data7.Recordset.Fields(18) = DataCombo1(17).Text
Data7.Recordset.Fields(21) = Text3(0).Text
Data7.Recordset.Update
Data7.Refresh
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(16).Enabled = False
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text5(i).Text = ""
Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
Data7.Recordset.Delete
Data7.Refresh
For i = 0 To 3
DataCombo1(i).Enabled = False
Next

DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text5(i).Text = ""
Next

DataCombo1(0).SetFocus
End Sub


Private Sub Command6_Click()
On Error Resume Next
Data7.Database.Execute "delete * from pldd where 料单编号='" & DataCombo1(2).Text & "'"
Data7.Database.Execute "INSERT INTO pldd SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "'"
Formd331.Data13.Refresh
Formd331.VSFlexGrid1.ColWidth(0) = 400
Formd331.VSFlexGrid1.ColWidth(1) = 0
Formd331.VSFlexGrid1.ColWidth(2) = 0
Formd331.VSFlexGrid1.ColWidth(5) = 400
Formd331.VSFlexGrid1.ColWidth(7) = 2000
Formd331.VSFlexGrid1.ColWidth(8) = 800
Formd331.VSFlexGrid1.ColWidth(10) = 600
Formd331.VSFlexGrid1.ColWidth(13) = 0
Formd331.VSFlexGrid1.ColWidth(14) = 0
Formd331.VSFlexGrid1.ColWidth(17) = 0
Formd331.VSFlexGrid1.ColWidth(18) = 0
Formd331.VSFlexGrid1.ColWidth(19) = 2600
Formd331.VSFlexGrid1.ColWidth(20) = 0
Formd331.VSFlexGrid1.ColWidth(22) = 0
Formd331.VSFlexGrid1.ColWidth(23) = 0
Formd331.VSFlexGrid1.ColWidth(24) = 0
Formd331.VSFlexGrid1.ColWidth(25) = 0
Unload Me
End Sub

Private Sub Command7_Click()
'On Error Resume Next
If MsgBox("按照模板 " + DataCombo7 + " 生成配料单吗？", vbYesNo) = vbNo Then Exit Sub
If DataCombo7 = "" Then
MsgBox ("请选择模板!")
Exit Sub
End If
Adodc2.RecordSource = "select * from CGGYMB where 模板编号='" & DataCombo7 & "'"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
Adodc2.Recordset.MoveFirst
Do While Not Adodc2.Recordset.EOF
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Data7.Database.Execute "insert into pldb(料单编号,工序名称,浴比,染化助库,染化助名称,配方单位,配方,校正值,次序号,批次,车速) values('" & DataCombo1(2) & "','" & Adodc2.Recordset.Fields(0) & "','" & DataCombo1(5) & "','" & Adodc2.Recordset.Fields(2) & "','" & Adodc2.Recordset.Fields(4) & "','" & Adodc2.Recordset.Fields(5) & "','" & Adodc2.Recordset.Fields(6) & "','1','" & Adodc2.Recordset.Fields(7) & "','','" & Adodc2.Recordset.Fields(8) & "')"

Adodc2.Recordset.MoveNext
Loop
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Data7.RecordSource = "SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "' ORDER BY val(工序名称),次序号"
Data7.Refresh

End Sub

Private Sub Command8_Click()
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
       Data7.RecordSource = "SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "' ORDER BY val(工序名称),次序号"
       Data7.Refresh
       Adodc8.Refresh
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY where '" & DataCombo1(4).Text & "' like 工艺名称 GROUP BY 工艺编号"
       Adodc12.Refresh
       Case 2
       Data7.RecordSource = "SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "' ORDER BY val(工序名称),次序号"
       Data7.Refresh

       If Data7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
       DataCombo1(i).Text = Data7.Recordset.Fields(i)
       Next
       DataCombo1(14).Text = Data7.Recordset.Fields(14)

       Case 6
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM rhzh where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       Case 13
       If DataCombo1(13).Text = "" Then
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh  where 染化助库名='" & DataCombo1(6).Text & "' AND 标志 like '%'+'" & DataCombo1(13).Text & "'+'%' GROUP BY 染料名称"
       Adodc8.Refresh
       End If
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY where '" & DataCombo1(4).Text & "' like 工艺名称 GROUP BY 工艺编号"
       Adodc12.Refresh

       Case 2
       Data7.RecordSource = "SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "'ORDER BY val(工序名称),次序号"
       Data7.Refresh

       If Data7.Recordset.EOF Then
        DataCombo1(16).Text = 1
       Else
         DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
       DataCombo1(i).Text = Data7.Recordset.Fields(i)
       Next
       DataCombo1(14).Text = Data7.Recordset.Fields(14)
       
       Case 6
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh   where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称"
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM rhzh where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       
       If InStr(DataCombo1(6), "染料") > 0 Then
       For i = 0 To 6
       DataCombo3(i).Text = "%"
       Next
       Else
       For i = 0 To 6
       DataCombo3(i).Text = "g/l"
       Next
       End If
       
       Case 13
       If DataCombo1(13).Text = "" Then
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh  where 染化助库名='" & DataCombo1(6).Text & "' AND 标志 like '%'+'" & DataCombo1(13).Text & "'+'%' GROUP BY 染料名称"
       Adodc8.Refresh
       End If
End Select
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode

End Sub


Private Sub DataCombo4_Click(Area As Integer)
On Error Resume Next
For i = 0 To 6
DataCombo2(i).Text = ""
Text1(i).Text = ""
Next
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT * FROM CGGY WHERE '" & DataCombo1(4).Text & "' like 工艺名称 AND  工艺编号='" & DataCombo4.Text & "' ORDER BY 序号"
Adodc13.Refresh
If Adodc13.Recordset.EOF Then
For i = 0 To 6
Text1(i).Text = ""
Next
Else
Adodc13.Recordset.MoveFirst
i = 0
Do While Not Adodc13.Recordset.EOF
DataCombo1(6).Text = Adodc13.Recordset.Fields(2)
DataCombo1(13).Text = Adodc13.Recordset.Fields(3)
DataCombo2(i).Text = Adodc13.Recordset.Fields(4)
DataCombo3(i).Text = Adodc13.Recordset.Fields(5)
Text1(i).Text = Adodc13.Recordset.Fields(6)
Text4(i).Text = Adodc13.Recordset.Fields(8)
i = i + 1
Adodc13.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Form_Load()

On Error Resume Next
Dim L As String

Data6.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data6.RecordSource = "pldb"
Data6.Refresh


For i = 0 To Data6.Recordset.Fields.count - 1
DataCombo1(i) = ""
Next

For i = 0 To 3
DataCombo1(i).Enabled = False
Next

DataCombo4.Text = ""

For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = 1
Text3(i).Text = ""
Text4(i).Text = ""
Text5(i).Text = ""
Next

Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

DataCombo1(10).Text = 1
DataCombo1(11).Text = Date
DataCombo1(11).Enabled = False
DataCombo1(11).Enabled = False
DataCombo1(14).Enabled = False
DataCombo1(15).Enabled = False
DataCombo1(15).Text = "大货"
DataCombo7 = ""
Data33.DatabaseName = App.Path & "\AccessBase\DB.mdb"

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select DISTINCT 模板编号 from CGGYMB ORDER by 模板编号"
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select 编号,工艺工序 from gx group by 编号,工艺工序 ORDER BY 工艺工序"
Adodc4.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select dw,IP from dw group by dw,IP ORDER BY IP"
Adodc5.Refresh



Data7.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data7.RecordSource = "SELECT * FROM pldb WHERE 料单编号='" & DataCombo1(2).Text & "' ORDER BY val(工序名称),次序号"
Data7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT 染化助库名 FROM rhzh GROUP BY 染化助库名"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT 标志 FROM rhzh GROUP BY 标志"
Adodc10.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY GROUP BY 工艺编号"
Adodc12.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

If Data7.Recordset.EOF Then
DataCombo1(16).Text = 1
Else
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
End If

DataCombo1(0).TabIndex = 0

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(5) = 400
VSFlexGrid1.ColWidth(7) = 2000
VSFlexGrid1.ColWidth(8) = 800
VSFlexGrid1.ColWidth(10) = 600
VSFlexGrid1.ColWidth(13) = 0
VSFlexGrid1.ColWidth(14) = 0
VSFlexGrid1.ColWidth(17) = 0
VSFlexGrid1.ColWidth(18) = 0
VSFlexGrid1.ColWidth(19) = 2600
VSFlexGrid1.ColWidth(20) = 0
VSFlexGrid1.ColWidth(22) = 0
VSFlexGrid1.ColWidth(23) = 0
VSFlexGrid1.ColWidth(24) = 0
VSFlexGrid1.ColWidth(25) = 0

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 16
       DataCombo1(16).Enabled = True
       Case 12
       DataCombo1(12).Enabled = True
       Case 11
       DataCombo1(10).Enabled = True
       Case 8
       DataCombo1(11).Enabled = True
       Case 9
       DataCombo1(12).Enabled = True
End Select
End Sub

Private Sub Text5_Change(Index As Integer)
Select Case Index
       Case Index
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT 染料名称 FROM v_rhzh where 简码 like '%'+'" & Text5(Index) & "'+'%' and 染化助库名='" & DataCombo1(6) & "' and 标志='用'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
DataCombo2(Index) = Adodc8.Recordset.Fields(0)
Else
DataCombo2(Index) = ""
End If
End Select
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Data7.Recordset.MoveFirst
Data7.Recordset.Move rs - 1
DataCombo1(4).Text = Data7.Recordset.Fields(3)
DataCombo1(6).Text = Data7.Recordset.Fields(5)
DataCombo2(0).Text = Data7.Recordset.Fields(6)
DataCombo3(0).Text = Data7.Recordset.Fields(7)
Text1(0).Text = Data7.Recordset.Fields(8)
Text2(0).Text = Data7.Recordset.Fields(9)
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
       Case Index
If Val(Text1(Index).Text) = 0 Then Text1(Index).Text = ""
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
       Case Index
       If Val(Text1(Index).Text) = 0 Then Text1(Index).Text = ""
       End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call MSF
End If
End Sub
Private Sub MSF()
On Error Resume Next
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
    If c = 9 Or c = 10 Or c = 16 Or c = 19 Then
    
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111.Text = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
   End If
End With
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
If c = 10 Then
If Val(Combo1111.Text) > 1 Then
If MsgBox("校正值大约用量的1倍以上，请确认是否继续？", vbYesNo) = vbNo Then Exit Sub
End If
End If
Data7.Recordset.MoveFirst
Data7.Recordset.Move r - 1
Data7.Recordset.Edit
Data7.Recordset.Fields(c - 1) = Combo1111.Text
Data7.Recordset.Update
VSFlexGrid1.Text = Combo1111.Text
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End If

If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
End Sub
