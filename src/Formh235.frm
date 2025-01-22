VERSION 5.00
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Formh235 
   BackColor       =   &H00C0E0FF&
   Caption         =   "确认工艺信息"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form23"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1111 
      Height          =   300
      Left            =   4560
      Style           =   1  'Simple Combo
      TabIndex        =   117
      Text            =   "Combo1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "模板确认"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   3840
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   330
      Left            =   1080
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Adodc16"
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
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   1920
      Top             =   9600
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
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc15"
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
      Left            =   3120
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Adodc14"
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
   Begin VB.Data Data7 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data6 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data15 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data Data14 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDataListLib.DataCombo DataCombo5 
      Height          =   330
      Left            =   12360
      TabIndex        =   92
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo5"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "配方单"
      Height          =   3495
      Left            =   3480
      TabIndex        =   24
      Top             =   360
      Width           =   11535
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   6
         Left            =   2880
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   5
         Left            =   2880
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   4
         Left            =   2880
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   3
         Left            =   2880
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   2
         Left            =   2880
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Index           =   0
         Left            =   2880
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   720
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "Formh235.frx":0000
         Height          =   330
         Left            =   1200
         TabIndex        =   91
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺编号"
         Text            =   "DataCombo4"
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   6
         Left            =   5800
         TabIndex        =   51
         Text            =   "Text3"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   5
         Left            =   5800
         TabIndex        =   50
         Text            =   "Text3"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   4
         Left            =   5800
         TabIndex        =   49
         Text            =   "Text3"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   3
         Left            =   5800
         TabIndex        =   48
         Text            =   "Text3"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   2
         Left            =   5800
         TabIndex        =   47
         Text            =   "Text3"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   1
         Left            =   5800
         TabIndex        =   46
         Text            =   "Text3"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Index           =   0
         Left            =   5800
         TabIndex        =   45
         Text            =   "Text3"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   8640
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   8640
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   8640
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   8640
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   8640
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   8640
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   8640
         TabIndex        =   38
         Text            =   "Text2"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   6
         Left            =   7200
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   7200
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   4
         Left            =   7200
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   3
         Left            =   7200
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   1880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   7200
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   1
         Left            =   7200
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   0
         Left            =   7200
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":0016
         Height          =   330
         Index           =   0
         Left            =   6480
         TabIndex        =   52
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh235.frx":002B
         Height          =   330
         Index           =   4
         Left            =   1200
         TabIndex        =   53
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺工序"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   5
         Left            =   1200
         TabIndex        =   54
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh235.frx":0040
         Height          =   330
         Index           =   6
         Left            =   1200
         TabIndex        =   55
         Top             =   2040
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "染化助库名"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh235.frx":0055
         Height          =   330
         Index           =   13
         Left            =   1200
         TabIndex        =   56
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "标志"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Formh235.frx":006B
         Height          =   330
         Index           =   12
         Left            =   1200
         TabIndex        =   57
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "工艺编号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   19
         Left            =   2880
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   11
         Left            =   9480
         TabIndex        =   59
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   16
         Left            =   9480
         TabIndex        =   60
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   17
         Left            =   9480
         TabIndex        =   61
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   18
         Left            =   9480
         TabIndex        =   62
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":0080
         Height          =   330
         Index           =   0
         Left            =   3600
         TabIndex        =   63
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":0095
         Height          =   330
         Index           =   1
         Left            =   3600
         TabIndex        =   64
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":00AA
         Height          =   330
         Index           =   2
         Left            =   3600
         TabIndex        =   65
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":00BF
         Height          =   330
         Index           =   3
         Left            =   3600
         TabIndex        =   66
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":00D4
         Height          =   330
         Index           =   4
         Left            =   3600
         TabIndex        =   67
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":00E9
         Height          =   330
         Index           =   5
         Left            =   3600
         TabIndex        =   68
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Formh235.frx":00FE
         Height          =   330
         Index           =   6
         Left            =   3600
         TabIndex        =   69
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "染料名称"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":0113
         Height          =   330
         Index           =   1
         Left            =   6480
         TabIndex        =   70
         Top             =   1080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":0128
         Height          =   330
         Index           =   2
         Left            =   6480
         TabIndex        =   71
         Top             =   1560
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":013D
         Height          =   330
         Index           =   3
         Left            =   6480
         TabIndex        =   72
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":0152
         Height          =   330
         Index           =   4
         Left            =   6480
         TabIndex        =   73
         Top             =   2280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":0167
         Height          =   330
         Index           =   5
         Left            =   6480
         TabIndex        =   74
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Formh235.frx":017C
         Height          =   330
         Index           =   6
         Left            =   6480
         TabIndex        =   90
         Top             =   3000
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         _Version        =   393216
         ListField       =   "dw"
         Text            =   "DataCombo3"
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Bindings        =   "Formh235.frx":0191
         Height          =   330
         Left            =   1200
         TabIndex        =   114
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
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
         TabIndex        =   115
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
         Index           =   21
         Left            =   2880
         TabIndex        =   113
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
         Left            =   5805
         TabIndex        =   89
         Top             =   240
         Width           =   615
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
         TabIndex        =   88
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
         Left            =   9480
         TabIndex        =   87
         Top             =   2400
         Width           =   1935
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
         Left            =   9480
         TabIndex        =   86
         Top             =   1680
         Width           =   1935
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
         Left            =   9480
         TabIndex        =   85
         Top             =   960
         Width           =   1935
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
         TabIndex        =   84
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
         TabIndex        =   83
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
         Left            =   3600
         TabIndex        =   82
         Top             =   240
         Width           =   2175
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
         Left            =   6480
         TabIndex        =   81
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
         Left            =   7200
         TabIndex        =   80
         Top             =   240
         Width           =   1335
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
         Left            =   9480
         TabIndex        =   79
         Top             =   240
         Width           =   1935
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
         TabIndex        =   78
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "校正值"
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
         Left            =   8640
         TabIndex        =   77
         Top             =   240
         Width           =   735
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
         TabIndex        =   76
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
         TabIndex        =   75
         Top             =   2040
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc13 
      Height          =   375
      Left            =   4800
      Top             =   9960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Adodc13"
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
      Left            =   4440
      Top             =   9960
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
      Caption         =   "Adodc12"
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
      Left            =   4080
      Top             =   9960
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
      Caption         =   "Adodc11"
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
      Left            =   3360
      Top             =   9960
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
      Caption         =   "Adodc10"
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
      Left            =   1800
      Top             =   9960
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
      Caption         =   "Adodc9"
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   375
      Left            =   2160
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Adodc8"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   2640
      Top             =   9960
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
      Left            =   2280
      Top             =   9960
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
      Left            =   1320
      Top             =   9960
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
      Left            =   720
      Top             =   9960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   1200
      Top             =   9960
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
      Left            =   1920
      Top             =   9960
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
      Left            =   2280
      Top             =   9960
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "客户信息"
      Height          =   3495
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   3255
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   0
         Left            =   1200
         TabIndex        =   18
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   14
         Left            =   1200
         TabIndex        =   22
         Top             =   2280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   330
         Index           =   15
         Left            =   1200
         TabIndex        =   23
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "颜色 "
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
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "色号"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "品名"
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "客户名称"
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
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "负责人"
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
         Index           =   14
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "生产类别"
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
         Index           =   15
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   11640
      Top             =   120
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      Height          =   375
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "浴比确认"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   2
      Left            =   7320
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   5
      Left            =   7320
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   6
      Left            =   7320
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "生成后退出"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3840
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   3840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   7
      Left            =   1800
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   8
      Left            =   1320
      TabIndex        =   95
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   9
      Left            =   1440
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   10
      Left            =   960
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   20
      Left            =   1800
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   21
      Left            =   1320
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   22
      Left            =   1440
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   23
      Left            =   960
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   24
      Left            =   1680
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   25
      Left            =   1200
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   26
      Left            =   1320
      TabIndex        =   104
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Index           =   27
      Left            =   840
      TabIndex        =   105
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formh235.frx":01A6
      Height          =   9255
      Left            =   120
      TabIndex        =   93
      Top             =   4440
      Width           =   18375
      _cx             =   32411
      _cy             =   16325
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
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "工艺曲线"
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
      Left            =   11520
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Formh235"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S1, S2 As Integer: Public c, r As Integer
Dim sz(56) As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub Command1_Click()
Data14.RecordSource = "select distinct 工序名称 from dpfda where 配方编号='" & DataCombo1(12).Text & "' order by 工序名称"
Data14.Refresh
If Data14.Recordset.EOF Then Exit Sub
If MsgBox("确定生成吗？", vbYesNo) = vbNo Then Exit Sub
Data14.Recordset.MoveFirst
i = 0
sz(i) = DataCombo1(0).Text
i = i + 1
sz(i) = DataCombo1(1).Text
i = i + 1
sz(i) = DataCombo1(2).Text
i = i + 1
sz(i) = DataCombo1(3).Text
i = i + 1
sz(i) = DataCombo1(12).Text
i = i + 1
sz(i) = DataCombo1(14).Text
i = i + 1
sz(i) = DataCombo1(11).Text
i = i + 1

Do While Not Data14.Recordset.EOF
Data15.RecordSource = "select 工序名称,浴比,染化助库,染化助名称,单位,配方,车速 from dpfda where 配方编号='" & DataCombo1(12).Text & "' and 工序名称='" & Data14.Recordset.Fields(0) & "' order by 次序号"
Data15.Refresh

If Not Data15.Recordset.EOF Then
Data15.Recordset.MoveFirst
Do While Not Data15.Recordset.EOF
sz(i) = Data15.Recordset.Fields(0) + "(" + Data15.Recordset.Fields(1) + ")" + Data15.Recordset.Fields(2) + "-" + Data15.Recordset.Fields(3) + "\" + Data15.Recordset.Fields(4) + "#" + Data15.Recordset.Fields(5) + "^" + Data15.Recordset.Fields(6)
i = i + 1
Data15.Recordset.MoveNext
Loop
End If

Data14.Recordset.MoveNext
Loop

If i < 57 Then
For L = i To 56
sz(L) = ""
Next
End If


Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
g_Cmd.CommandText = "dbdpfd('" & sz(0) & "','" & sz(1) & "','" & sz(2) & "','" & sz(3) & "','" & sz(4) & "','" & sz(5) & "','" & sz(6) & "','" & sz(7) & "','" & sz(8) & "','" & sz(9) & "','" & sz(10) & "','" & sz(11) & "','" & sz(12) & "','" & sz(13) & "','" & sz(14) & "','" & sz(15) & "','" & sz(16) & "','" & sz(17) & "','" & sz(18) & "','" & sz(19) & "','" & sz(20) & "','" & sz(21) & "','" & sz(22) & "','" & sz(23) & "','" & sz(24) & "','" & sz(25) & "','" & sz(26) & "','" & sz(27) & "','" & sz(28) & "','" & sz(29) & "','" & sz(30) & "','" & sz(31) & "','" & sz(32) & "','" & sz(33) & "','" & sz(34) & "','" & sz(35) & "','" & sz(36) & "','" & sz(37) & "','" & sz(38) & "','" & sz(39) & "','" & sz(40) & "','" & sz(41) & "','" & sz(42) & "','" & sz(43) & "','" & sz(44) & "','" & sz(45) & "','" & sz(46) & "','" & sz(47) & "','" & sz(48) & "','" & sz(49) & "','" & sz(50) & "','" & sz(51) & "','" & sz(52) & "','" & sz(53) & "','" & sz(54) & "','" & sz(55) & "','" & sz(56) & "')"
g_Cmd.Execute           ' 执行存储过程
g_Cmd.Cancel

sql1 = "update zh2 set rq='" & Now & "' where dh='" & DataCombo1(12).Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Unload Me

End Sub




Private Sub Command12_Click()
If DataCombo1(12).Text = "" Then
MsgBox ("没有配方编号")
Exit Sub
End If
If DataCombo1(5).Text = "" Then
MsgBox ("请输入浴比")
Exit Sub
End If
Data7.Database.Execute "UPDATE dpfda SET 浴比='" & DataCombo1(5).Text & "' WHERE 配方编号='" & DataCombo1(12).Text & "'"
Data7.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "'ORDER BY val(工序名称),次序号"
Data7.Refresh
       If Data7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
       End If
End Sub

Private Sub Command2_Click()
'On Error Resume Next
If DataCombo1(2).Text = "" Or DataCombo1(12).Text = "" Then
MsgBox ("色号配方编号须填完整！")
Exit Sub
End If

For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(10).Enabled = False
DataCombo1(11).Enabled = False
DataCombo1(12).Enabled = False

DataCombo1(13).Text = ""        '''''''''''''代码清离

For i = 0 To 6     '''''''''''''''''''''''''
If Text1(i).Text <> "" Then
DataCombo1(7).Text = DataCombo2(i).Text
DataCombo1(8).Text = DataCombo3(i).Text
DataCombo1(9).Text = Text1(i).Text
DataCombo1(10).Text = Text2(i).Text
DataCombo1(17).Text = Text4(i).Text
DataCombo1(19).Text = Text3(i).Text
Data6.Recordset.AddNew
For p = 0 To Data6.Recordset.Fields.count - 1
Data6.Recordset.Fields(p) = DataCombo1(p).Text
Next
Data6.Recordset.Fields(16) = Data7.Recordset.RecordCount + 1
Data6.Recordset.Update
Data7.Refresh
End If
Next
                '''''''''''''''''''''''
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = ""
Text3(i).Text = ""
Text4(i).Text = ""
Next
DataCombo1(16).Enabled = False
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
DataCombo1(4).SetFocus
End Sub

Private Sub Command3_Click()
If DataCombo1(2).Text = "" Or DataCombo1(12).Text = "" Then
MsgBox ("色号配方编号须填完整！")
Exit Sub
End If

Data7.Recordset.Edit
DataCombo1(7).Text = DataCombo2(0).Text
DataCombo1(8).Text = DataCombo3(0).Text
DataCombo1(9).Text = Text1(0).Text
DataCombo1(10).Text = Text2(0).Text
DataCombo1(17).Text = Text4(0).Text
DataCombo1(19).Text = Text3(0).Text
For i = 0 To Data7.Recordset.Fields.count - 1
Data7.Recordset.Fields(i) = DataCombo1(i).Text
Next
Data7.Recordset.Update
Data7.Refresh
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo1(16).Enabled = False
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = ""
Text3(i).Text = ""
Text4(i).Text = ""
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
Text2(i).Text = ""
Text3(i).Text = ""
Text4(i).Text = ""
Next
DataCombo1(0).SetFocus
End Sub

Private Sub Command5_Click()
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
Data7.Database.Execute "insert into dpfda(浴比,加工单位,品名,色号,颜色,工序名称,染化助库,染化助名称,单位,配方,校正值,配方日期,配方编号,负责人,次序号,车速) values('" & DataCombo1(5) & "','" & DataCombo1(0) & "','" & DataCombo1(1) & "','" & DataCombo1(2) & "','" & DataCombo1(3) & "','" & Adodc2.Recordset.Fields(0) & "','" & Adodc2.Recordset.Fields(2) & "','" & Adodc2.Recordset.Fields(4) & "','" & Adodc2.Recordset.Fields(5) & "','" & Adodc2.Recordset.Fields(6) & "','1','" & DataCombo1(11) & "','" & DataCombo1(12) & "','" & DataCombo1(14) & "','" & Adodc2.Recordset.Fields(7) & "','" & Adodc2.Recordset.Fields(8) & "')"

Adodc2.Recordset.MoveNext
Loop
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Data7.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "' ORDER BY val(工序名称),次序号"
Data7.Refresh
End Sub

Private Sub Command6_Click()
Unload Me
End Sub


Private Sub Command8_Click()
If Data7.Recordset.EOF Then
DataCombo1(16).Text = 1
Else
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
End If
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY WHERE '" & DataCombo1(4).Text & "' like 工艺名称 GROUP BY 工艺编号"
       Adodc12.Refresh
       Case 12
       Data7.DatabaseName = App.Path & "\AccessBase\db.mdb"
       Data7.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "' ORDER BY val(工序名称),次序号"
       Data7.Refresh

       If Data7.Recordset.EOF Then
       DataCombo1(16).Text = 1
       Else
       DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
'       DataCombo1(i).Text = Data7.Recordset.Fields(i)
       Next
       Case 6
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
       Adodc10.Refresh
       Case 13
       If DataCombo1(13).Text = "" Then
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(6).Text & "' AND 标志 like '" & DataCombo1(13).Text & "' GROUP BY 染料名称"
       Adodc8.Refresh
       End If
End Select
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Select Case Index
       Case 4
       Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc12.RecordSource = "SELECT 工艺编号 FROM CGGY WHERE '" & DataCombo1(4).Text & "' like 工艺名称 GROUP BY 工艺编号"
       Adodc12.Refresh

       Case 12
       Data7.DatabaseName = App.Path & "\AccessBase\db.mdb"
       Data7.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "' ORDER BY val(工序名称),次序号"
       Data7.Refresh

       If Data7.Recordset.EOF Then
        DataCombo1(16).Text = 1
       Else
         DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
       End If
       
       For i = 0 To 3
       DataCombo1(i).Text = Data7.Recordset.Fields(i)
       Next
       Case 6
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称"
       Adodc8.Refresh
       Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc10.RecordSource = "SELECT 标志 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 标志 "
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
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH where 染化助库名='" & DataCombo1(6).Text & "' GROUP BY 染料名称 "
       Adodc8.Refresh
       Else
       Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc8.RecordSource = "SELECT 染料名称 FROM RHZH  where 染化助库名='" & DataCombo1(6).Text & "' AND 标志 like '" & DataCombo1(13).Text & "' GROUP BY 染料名称"
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
Adodc13.RecordSource = "SELECT * FROM CGGY WHERE 工艺编号='" & DataCombo4.Text & "' AND '" & DataCombo1(4).Text & "' like 工艺名称 ORDER BY 序号"
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

'On Error Resume Next
Dim L As String

mb = 1
Timer1.Enabled = False
ProgressBar1.Visible = False
For i = 0 To 3
DataCombo1(i).Enabled = False
Next
DataCombo4.Text = ""
DataCombo5.Text = ""
DataCombo7.Text = ""
For i = 0 To 6
DataCombo2(i).Text = ""
DataCombo3(i).Text = "%"
Text1(i).Text = ""
Text2(i).Text = ""
Text3(i).Text = ""
Text4(i).Text = ""
Text5(i).Text = ""
Next

Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

DataCombo1(10).Text = 1
DataCombo1(11).Text = Date
DataCombo1(11).Enabled = False
DataCombo1(11).Enabled = False
DataCombo1(12).Enabled = False
DataCombo1(14).Enabled = False
DataCombo1(15).Enabled = False
DataCombo1(16).Text = ""
DataCombo1(17).Text = ""

Data7.DatabaseName = App.Path & "\AccessBase\db.mdb"
Data7.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "' ORDER BY val(工序名称),次序号"
Data7.Refresh

For i = 0 To Data7.Recordset.Fields.count - 1
DataCombo1(i) = ""
Next

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


Data6.DatabaseName = App.Path & "\AccessBase\DB.mdb"
Data6.RecordSource = "SELECT * FROM dpfda WHERE 配方编号='" & DataCombo1(12).Text & "' ORDER BY val(工序名称),次序号"
Data6.Refresh


Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc8.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "SELECT 染化助库名 FROM RHZH GROUP BY 染化助库名"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT 标志 FROM RHZH GROUP BY 标志"
Adodc10.Refresh


Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT 工艺编号 FROM CGGY GROUP BY 工艺编号"
Adodc10.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Data14.DatabaseName = App.Path & "\AccessBase\db.mdb"
Data15.DatabaseName = App.Path & "\AccessBase\db.mdb"

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

If Data7.Recordset.EOF Then
DataCombo1(16).Text = 1
Else
DataCombo1(16).Text = Data7.Recordset.RecordCount + 1
End If

DataCombo1(0).TabIndex = 0

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(4) = 0
VSFlexGrid1.ColWidth(5) = 1000
VSFlexGrid1.ColWidth(6) = 400
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(8) = 2000
VSFlexGrid1.ColWidth(9) = 1000
VSFlexGrid1.ColWidth(10) = 1500
VSFlexGrid1.ColWidth(11) = 800
VSFlexGrid1.ColWidth(12) = 1200
VSFlexGrid1.ColWidth(15) = 0
VSFlexGrid1.ColWidth(16) = 0
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
Adodc8.RecordSource = "SELECT 染料名称 FROM rhzh where 简码 like '%'+'" & Text5(Index) & "'+'%' and 染化助库名='" & DataCombo1(6) & "' and 标志='用'"
Adodc8.Refresh
If Not Adodc8.Recordset.EOF Then
DataCombo2(Index) = Adodc8.Recordset.Fields(0)
Else
DataCombo2(Index) = ""
End If
End Select
End Sub

Private Sub VSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
S1 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
S2 = VSFlexGrid1.RowSel
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Data7.Recordset.MoveFirst
Data7.Recordset.Move rs - 1
For i = 0 To Data7.Recordset.Fields.count - 1
If i <> 14 Then
DataCombo1(i).Text = Data7.Recordset.Fields(i)
End If
Next
DataCombo2(0).Text = DataCombo1(7).Text
DataCombo3(0).Text = DataCombo1(8).Text
Text1(0).Text = DataCombo1(9).Text
Text2(0).Text = DataCombo1(10).Text
Text3(0).Text = DataCombo1(19).Text
Text4(0).Text = DataCombo1(17).Text
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

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub vSFlexGrid1_Dbl()
With VSFlexGrid1
    c = .col: r = .Row    '''''C列，，R行
If c = 9 Or c = 10 Or c = 11 Or c = 18 Or c = 20 Or c = 8 Or c = 17 Then
        Combo1111.Left = .Left + .ColPos(c)
        Combo1111.Top = .Top + .RowPos(r)
        Combo1111.Width = .ColWidth(c)
        Combo1111.Height = .RowHeight(r)
        Combo1111 = .Text
        Combo1111.Visible = True
        Combo1111.SetFocus
End If
End With
End Sub

Private Sub vSFlexGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Call vSFlexGrid1_Dbl
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
       Case Index
       If Val(Text1(Index).Text) = 0 Then Text1(Index).Text = ""
       End Select
End Sub

Private Sub combo1111_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyEscape Then
    Combo1111.Visible = False
    VSFlexGrid1.SetFocus
    Exit Sub
End If
If KeyAscii = vbKeyReturn Then

Data7.Recordset.MoveFirst
Data7.Recordset.Move r - 1
Data7.Recordset.Edit
If (c = 17 Or c = 10) Then
Data7.Recordset.Fields(c - 1) = Val(Combo1111.Text)
VSFlexGrid1.Text = Val(Combo1111.Text)
End If
If c = 8 Then
Adodc14.RecordSource = "select distinct 染料名称,染化助库名 from rhzh where 简码='" & Combo1111 & "'"
Adodc14.Refresh
If Not Adodc14.Recordset.EOF Then
Combo1111 = Adodc14.Recordset.Fields(0)
Data7.Recordset.Fields(c - 2) = Adodc14.Recordset.Fields(1)
VSFlexGrid1.TextMatrix(r, c - 1) = Adodc14.Recordset.Fields(1)
End If
Data7.Recordset.Fields(c - 1) = Combo1111.Text
VSFlexGrid1.Text = Combo1111.Text
End If
If c <> 8 And c <> 17 And c <> 10 Then
Data7.Recordset.Fields(c - 1) = Combo1111.Text
VSFlexGrid1.Text = Combo1111.Text
End If
Data7.Recordset.Update
Combo1111.Visible = False
VSFlexGrid1.SetFocus
End If
End Sub



