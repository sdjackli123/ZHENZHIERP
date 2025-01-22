VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Formw35 
   BackColor       =   &H00C0E0FF&
   Caption         =   "凭证制作"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form35"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   5280
      Width           =   975
   End
   Begin VB.Data Data14 
      Caption         =   "Data14"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   285
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw35.frx":0000
      Left            =   1320
      List            =   "Formw35.frx":000A
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12360
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2160
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   12600
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "增加"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "保存"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Formw35.frx":001E
      Left            =   12360
      List            =   "Formw35.frx":002E
      TabIndex        =   9
      Text            =   "Combo2"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data7 
      Caption         =   "Data7"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data8 
      Caption         =   "Data8"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data9 
      Caption         =   "Data9"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data10 
      Caption         =   "Data10"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data11 
      Caption         =   "Data11"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data12 
      Caption         =   "Data12"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data13 
      Caption         =   "Data13"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Text            =   "Text5"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formw35.frx":005A
      Height          =   330
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ListField       =   "MC"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Height          =   330
      Index           =   0
      Left            =   12600
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   0
      Left            =   10440
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   0
      Left            =   6000
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Index           =   0
      Left            =   3840
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw35.frx":006E
      Height          =   360
      Index           =   0
      Left            =   1560
      TabIndex        =   20
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "摘要"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Formw35.frx":0082
      Height          =   3615
      Left            =   360
      TabIndex        =   24
      Top             =   5880
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   14
      BackColorFixed  =   8421631
      BackColorBkg    =   32896
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   5280
      TabIndex        =   25
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   22806529
      CurrentDate     =   39883
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw35.frx":0096
      Height          =   360
      Index           =   1
      Left            =   1560
      TabIndex        =   26
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "摘要"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Index           =   1
      Left            =   3840
      TabIndex        =   27
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   1
      Left            =   6000
      TabIndex        =   28
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   1
      Left            =   8400
      TabIndex        =   29
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw35.frx":00AA
      Height          =   360
      Index           =   2
      Left            =   1560
      TabIndex        =   30
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "摘要"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Index           =   2
      Left            =   3840
      TabIndex        =   31
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   2
      Left            =   6000
      TabIndex        =   32
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   2
      Left            =   8400
      TabIndex        =   33
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw35.frx":00BE
      Height          =   360
      Index           =   3
      Left            =   1560
      TabIndex        =   34
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "摘要"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Index           =   3
      Left            =   3840
      TabIndex        =   35
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   3
      Left            =   6000
      TabIndex        =   36
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   3
      Left            =   8400
      TabIndex        =   37
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Formw35.frx":00D2
      Height          =   360
      Index           =   4
      Left            =   1560
      TabIndex        =   38
      Top             =   3720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ListField       =   "摘要"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Height          =   330
      Index           =   4
      Left            =   3840
      TabIndex        =   39
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Height          =   330
      Index           =   4
      Left            =   6000
      TabIndex        =   40
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo4 
      Height          =   330
      Index           =   4
      Left            =   8400
      TabIndex        =   41
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo4"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   1
      Left            =   10440
      TabIndex        =   42
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   2
      Left            =   10440
      TabIndex        =   43
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   3
      Left            =   10440
      TabIndex        =   44
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo5 
      Height          =   330
      Index           =   4
      Left            =   10440
      TabIndex        =   45
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ListField       =   ""
      Text            =   "DBCombo5"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Height          =   330
      Index           =   1
      Left            =   12600
      TabIndex        =   46
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Height          =   330
      Index           =   2
      Left            =   12600
      TabIndex        =   47
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Height          =   330
      Index           =   3
      Left            =   12600
      TabIndex        =   48
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo6 
      Height          =   330
      Index           =   4
      Left            =   12600
      TabIndex        =   49
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Text            =   "DBCombo6"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formw35.frx":00E6
      Height          =   330
      Index           =   1
      Left            =   4680
      TabIndex        =   50
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ListField       =   "MC"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formw35.frx":00FA
      Height          =   330
      Index           =   2
      Left            =   7320
      TabIndex        =   51
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ListField       =   "MC"
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formw35.frx":010E
      Height          =   330
      Index           =   3
      Left            =   9840
      TabIndex        =   52
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ListField       =   "MC "
      Text            =   "DBCombo7"
   End
   Begin MSDBCtls.DBCombo DBCombo7 
      Bindings        =   "Formw35.frx":0122
      Height          =   330
      Index           =   4
      Left            =   12840
      TabIndex        =   53
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ListField       =   "MC"
      Text            =   "DBCombo7"
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label11"
      Height          =   255
      Index           =   4
      Left            =   12120
      TabIndex        =   97
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label11"
      Height          =   255
      Index           =   3
      Left            =   12120
      TabIndex        =   96
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label11"
      Height          =   255
      Index           =   2
      Left            =   12120
      TabIndex        =   95
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label11"
      Height          =   255
      Index           =   1
      Left            =   12120
      TabIndex        =   94
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label11"
      Height          =   255
      Index           =   0
      Left            =   12120
      TabIndex        =   93
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label10"
      Height          =   255
      Index           =   4
      Left            =   9960
      TabIndex        =   92
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label10"
      Height          =   255
      Index           =   3
      Left            =   9960
      TabIndex        =   91
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label10"
      Height          =   255
      Index           =   2
      Left            =   9960
      TabIndex        =   90
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label10"
      Height          =   255
      Index           =   1
      Left            =   9960
      TabIndex        =   89
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label10"
      Height          =   255
      Index           =   0
      Left            =   9960
      TabIndex        =   88
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label9"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   87
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label9"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   86
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label9"
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   85
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label9"
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   84
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label9"
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   83
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   82
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   81
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   80
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   79
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Label8"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   78
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " 记 账 凭 证"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   5880
      TabIndex        =   76
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   14880
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   14640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   240
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   14880
      X2              =   14880
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   240
      X2              =   14640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "贷方科目："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   75
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   14520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      X1              =   240
      X2              =   14640
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line8 
      X1              =   240
      X2              =   14520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   74
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "操作月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   73
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "银付子第"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11280
      TabIndex        =   72
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "固定号            方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   6
      Left            =   14880
      TabIndex        =   68
      Top             =   1440
      Width           =   255
   End
   Begin VB.Line Line9 
      Index           =   0
      X1              =   3840
      X2              =   8040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "总 账 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   67
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " 明 细 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   66
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Line Line10 
      X1              =   5880
      X2              =   5880
      Y1              =   1800
      Y2              =   4080
   End
   Begin VB.Line Line11 
      X1              =   3720
      X2              =   3720
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line12 
      X1              =   8280
      X2              =   8280
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   240
      X2              =   14880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line13 
      X1              =   5640
      X2              =   9840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line14 
      X1              =   240
      X2              =   14880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "附原始单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   65
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "会计主管"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   64
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "记账"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   63
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "复核"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   6840
      TabIndex        =   62
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出纳"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   9360
      TabIndex        =   61
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "制单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   12360
      TabIndex        =   60
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "凭证类型："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   13
      Left            =   11280
      TabIndex        =   59
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "明 细 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   58
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "总 账 科 目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   57
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Line Line9 
      Index           =   1
      X1              =   8280
      X2              =   12480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line15 
      X1              =   12480
      X2              =   12480
      Y1              =   1440
      Y2              =   4560
   End
   Begin VB.Line Line16 
      X1              =   10320
      X2              =   10320
      Y1              =   1800
      Y2              =   4080
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      TabIndex        =   56
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "原始单据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   15
      Left            =   240
      TabIndex        =   55
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "合     计"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   10920
      TabIndex        =   54
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "借  方  科  目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   3
      Left            =   3840
      TabIndex        =   70
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "摘           要"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   1
      Left            =   1560
      TabIndex        =   71
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "贷  方  科  目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   14
      Left            =   8400
      TabIndex        =   77
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  金          额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   5
      Left            =   12480
      TabIndex        =   69
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "Formw35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public K1, K2, PZH As String: Dim SZSZ(4) As Integer
Private Sub Combo1_Click()
On Error Resume Next
If Combo2.Text = "付款凭证" Then

If Combo1.Text = "现金" Then
For i = 0 To 4
DBCombo4(i).Text = Combo1.Text
Next

Label7.Caption = "2"
Data10.RecordSource = "SELECT MAX(VAL(MID(CLFKPZ.凭证号,3))) FROM CLFKPZ WHERE INSTR(CLFKPZ.凭证号,'2-')>0 AND CLFKPZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "2-1"
If Not Data10.Recordset.EOF Then
PZH = "2-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Text2.Text = PZH
End If
If Combo1.Text = "银行存款" Then
For i = 0 To 4
DBCombo4(i).Text = Combo1.Text
Next

Label7.Caption = "4"
Data10.RecordSource = "SELECT MAX(VAL(MID(CLFKPZ.凭证号,3))) FROM CLFKPZ  WHERE INSTR(CLFKPZ.凭证号,'4-')>0 AND CLFKPZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "4-1"
If Not Data10.Recordset.EOF Then
PZH = "4-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Text2.Text = PZH
End If

End If

If Combo2.Text = "收款凭证" Then
For i = 0 To 4
DBCombo2(i).Text = Combo1.Text
Next

If Combo1.Text = "现金" Then
Label7.Caption = "1"
Data10.RecordSource = "SELECT MAX(VAL(MID(CLSKPZ.凭证号,3))) FROM CLSKPZ  WHERE INSTR(CLSKPZ.凭证号,'1-')>0 AND CLSKPZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "1-1"
If Not Data10.Recordset.EOF Then
PZH = "1-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Text2.Text = PZH
End If
If Combo1.Text = "银行存款" Then
For i = 0 To 4
DBCombo2(i).Text = Combo1.Text
Next

Label7.Caption = "3"
Data10.RecordSource = "SELECT MAX(VAL(MID(CLSKPZ.凭证号,3))) FROM CLSKPZ  WHERE INSTR(CLSKPZ.凭证号,'3-')>0 AND CLSKPZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "3-1"
If Not Data10.Recordset.EOF Then
PZH = "3-" + Trim(Data10.Recordset.Fields(0) + 1)
End If
Text2.Text = PZH
End If

End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
If Combo2.Text = "转账凭证" Then
Text2.Text = ""
For i = 0 To 4
DBCombo2(i).Enabled = True
DBCombo2(i).Text = ""
DBCombo3(i).Enabled = False
DBCombo3(i).Text = ""
DBCombo4(i).Enabled = True
DBCombo4(i).Text = ""
DBCombo5(i).Enabled = True
DBCombo5(i).Text = ""
Label9(i).Enabled = True
Label10(i).Enabled = True
Label11(i).Enabled = True
Label8(i).Enabled = True
Next
Label1(2).Caption = ""
Combo1.Visible = False
Label7.Caption = "5"

Data10.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.凭证号,3))) FROM CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "5-1"
If Not Data10.Recordset.EOF Then
PZH = "5-" + Trim(Data10.Recordset.Fields(0) + 1)
Text2.Text = PZH
End If

End If

If Combo2.Text = "付款凭证" Then
Text2.Text = ""
Label1(2).Caption = "贷方总账科目"
Combo1.Visible = True
For i = 0 To 4
DBCombo4(i).Text = ""
DBCombo4(i).Enabled = False
DBCombo5(i).Text = ""
DBCombo5(i).Enabled = False
DBCombo2(i).Text = ""
DBCombo2(i).Enabled = True
DBCombo3(i).Text = ""
DBCombo3(i).Enabled = True
'Label10(i).Enabled = False
'Label11(i).Enabled = False
Label8(i).Enabled = True
Label9(i).Enabled = True
Next
End If

If Combo2.Text = "收款凭证" Then
Text2.Text = ""
Label1(2).Caption = "借方总账科目"
Combo1.Visible = True
For i = 0 To 4
DBCombo4(i).Text = ""
DBCombo4(i).Enabled = True
DBCombo5(i).Text = ""
DBCombo5(i).Enabled = True
DBCombo2(i).Text = ""
DBCombo2(i).Enabled = False
DBCombo3(i).Text = ""
DBCombo3(i).Enabled = False
Label10(i).Enabled = True
Label11(i).Enabled = True
'Label8(i).Enabled = False
'Label9(i).Enabled = False

Next
End If

If Combo2.Text = "成本凭证" Then
Text2.Text = ""
For i = 0 To 4
DBCombo2(i).Enabled = True
DBCombo2(i).Text = ""
DBCombo3(i).Enabled = False
DBCombo3(i).Text = ""
DBCombo4(i).Enabled = True
DBCombo4(i).Text = ""
DBCombo5(i).Enabled = True
DBCombo5(i).Text = ""
Label9(i).Enabled = True
Label10(i).Enabled = True
Label11(i).Enabled = True
Label8(i).Enabled = True
Next
Label1(2).Caption = ""
Combo1.Visible = False
Label7.Caption = "S"

Data10.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.凭证号,3))) FROM CLSCCB WHERE CLSCCB.日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "')"
Data10.Refresh
PZH = "S-1"
If Not Data10.Recordset.EOF Then
PZH = "S-" + Trim(Data10.Recordset.Fields(0) + 1)
Text2.Text = PZH
End If

End If

Combo1.Text = ""
End Sub

Private Sub Command1_Click()
For i = 0 To 4
DBCombo1(i).Text = ""
DBCombo2(i).Text = ""
DBCombo3(i).Text = ""
DBCombo4(i).Text = ""
DBCombo5(i).Text = ""
DBCombo6(i).Text = ""
DBCombo7(i).Text = ""
Next
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Combo2.Text = "转账凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLZZPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
For i = 0 To 4
If DBCombo6(i).Text <> "" Then
If DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DBCombo1(i).Text
Data2.Recordset.Fields(1) = DBCombo2(i).Text
Data2.Recordset.Fields(2) = DBCombo3(i).Text
Data2.Recordset.Fields(3) = DBCombo4(i).Text
Data2.Recordset.Fields(4) = DBCombo5(i).Text
Data2.Recordset.Fields(5) = DBCombo6(i).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(i).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
End If
End If
Next
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLZZPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If
'''''''''''''''''''''''''''''''''''''''''''''

If Combo2.Text = "付款凭证" Then
If Combo1.Text = "" Or Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Data2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLFKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
For i = 0 To 4
If DBCombo6(i).Text <> "" Then
If DBCombo2(i).Text = "" Or DBCombo6(i).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DBCombo1(i).Text
Data2.Recordset.Fields(1) = DBCombo2(i).Text
Data2.Recordset.Fields(2) = DBCombo3(i).Text
Data2.Recordset.Fields(3) = DBCombo4(i).Text
Data2.Recordset.Fields(4) = DBCombo5(i).Text
Data2.Recordset.Fields(5) = DBCombo6(i).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(i).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
End If
End If
Next
Data2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLFKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh

End If

If Combo2.Text = "成本凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，1确认！")
Exit Sub
End If
Data2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSCCB.凭证号='" & Text2.Text & "'"
Data2.Refresh
For i = 0 To 4
If DBCombo6(i).Text <> "" Then
If DBCombo2(i).Text = "" Or DBCombo4(i).Text = "" Or DBCombo6(i).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，2确认！")
Exit Sub
Else
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DBCombo1(i).Text
Data2.Recordset.Fields(1) = DBCombo2(i).Text
Data2.Recordset.Fields(2) = DBCombo3(i).Text
Data2.Recordset.Fields(3) = DBCombo4(i).Text
Data2.Recordset.Fields(4) = DBCombo5(i).Text
Data2.Recordset.Fields(5) = DBCombo6(i).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(i).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
End If
End If
Next
Data2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSCCB.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

If Combo2.Text = "收款凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
Data2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
For i = 0 To 4
If DBCombo6(i).Text <> "" Then
If DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DBCombo1(i).Text
Data2.Recordset.Fields(1) = DBCombo2(i).Text
Data2.Recordset.Fields(2) = DBCombo3(i).Text
Data2.Recordset.Fields(3) = DBCombo4(i).Text
Data2.Recordset.Fields(4) = DBCombo5(i).Text
Data2.Recordset.Fields(5) = DBCombo6(i).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(i).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
End If
End If
Next
Data2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

End Sub

Private Sub Command4_Click()
If Combo2.Text = "转账凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If DBCombo1(0).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo1(0).Text
Data2.Recordset.Fields(1) = DBCombo2(0).Text
Data2.Recordset.Fields(2) = DBCombo3(0).Text
Data2.Recordset.Fields(3) = DBCombo4(0).Text
Data2.Recordset.Fields(4) = DBCombo5(0).Text
Data2.Recordset.Fields(5) = DBCombo6(0).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(0).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLZZPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
MsgBox ("修改成功！")
End If
End If
'''''''''''''''''''''''''''''''''''''''''''''

If Combo2.Text = "付款凭证" Then
If Combo1.Text = "" Or Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If DBCombo1(0).Text = "" Or DBCombo2(0).Text = "" Or DBCombo3(0).Text = "" Or DBCombo6(0).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo1(0).Text
Data2.Recordset.Fields(1) = DBCombo2(0).Text
Data2.Recordset.Fields(2) = DBCombo3(0).Text
Data2.Recordset.Fields(3) = DBCombo4(0).Text
Data2.Recordset.Fields(4) = DBCombo5(0).Text
Data2.Recordset.Fields(5) = DBCombo6(0).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(0).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
Data2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLFKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
MsgBox ("修改成功！")
End If
End If

If Combo2.Text = "收款凭证" Then
If Combo1.Text = "" Or Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，清确认！")
Exit Sub
End If
If DBCombo1(0).Text = "" Or DBCombo2(0).Text = "" Or DBCombo5(0).Text = "" Or DBCombo6(0).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo1(0).Text
Data2.Recordset.Fields(1) = DBCombo2(0).Text
Data2.Recordset.Fields(2) = DBCombo3(0).Text
Data2.Recordset.Fields(3) = DBCombo4(0).Text
Data2.Recordset.Fields(4) = DBCombo5(0).Text
Data2.Recordset.Fields(5) = DBCombo6(0).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(0).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
Data2.RecordSource = "SELECT * FROM CLSKPZ WHERE  日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
MsgBox ("修改成功！")
End If
End If


If Combo2.Text = "成本凭证" Then
If Combo2.Text = "" Or Text2.Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
End If
If DBCombo1(0).Text = "" Or DBCombo7(4).Text = "" Then
MsgBox ("输入有误，确认！")
Exit Sub
Else
Data2.Recordset.Edit
Data2.Recordset.Fields(0) = DBCombo1(0).Text
Data2.Recordset.Fields(1) = DBCombo2(0).Text
Data2.Recordset.Fields(2) = DBCombo3(0).Text
Data2.Recordset.Fields(3) = DBCombo4(0).Text
Data2.Recordset.Fields(4) = DBCombo5(0).Text
Data2.Recordset.Fields(5) = DBCombo6(0).Text
Data2.Recordset.Fields(6) = Text2.Text
Data2.Recordset.Fields(7) = DTPicker3.Value
Data2.Recordset.Fields(8) = Text5(0).Text
Data2.Recordset.Fields(9) = DBCombo7(1).Text
Data2.Recordset.Fields(10) = DBCombo7(2).Text
Data2.Recordset.Fields(11) = DBCombo7(4).Text
Data2.Recordset.Fields(12) = Text3.Text
Data2.Recordset.Update
Data2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSCCB.凭证号='" & Text2.Text & "'"
Data2.Refresh
MsgBox ("修改成功！")
End If
End If

End Sub

Private Sub Command5_Click()
If MsgBox("删除将不能恢复！", vbYesNo) = vbNo Then Exit Sub
If Combo2.Text = "转账凭证" Then
Data2.Recordset.Delete
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND 凭证号='" & Text2.Text & "'"
Data2.Refresh
End If
'''''''''''''''''''''''''''''''''''''''''''''
If Combo2.Text = "付款凭证" Then
Data2.Recordset.Delete
Data2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND 凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

If Combo2.Text = "收款凭证" Then
Data2.Recordset.Delete
Data2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND 凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

If Combo2.Text = "成本凭证" Then
Data2.Recordset.Delete
Data2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND 凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

End Sub

Private Sub Command6_Click()
Data1.Refresh
Data5.Refresh
Data6.Refresh
Data7.Refresh
Data8.Refresh
Data9.Refresh
End Sub

Private Sub Command7_Click()
If Combo2.Text = "" Or Text2.Text = "" Then
MsgBox ("请输入凭证类别和凭证号")
Exit Sub
End If
If Data1.Recordset.EOF Then Exit Sub
Call PZDY(Combo2.Text, Text2.Text)
End Sub

Private Sub DTPicker3_Change()
Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text1.Text = Data14.Recordset.Fields(2)
End If
End Sub

Private Sub DTPicker3_CloseUp()
Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text1.Text = Data14.Recordset.Fields(2)
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Combo1.Text = ""
DTPicker3.Value = Date

Data14.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CJBB.mdb"
Data14.RecordSource = "select * from RQSD where cdate('" & DTPicker3.Value & "') between 起始日期 and 结束日期"
Data14.Refresh
If Data14.Recordset.EOF Then
MsgBox ("期间有误")
Else
K1 = Data14.Recordset.Fields(0)
K2 = Data14.Recordset.Fields(1)
Text1.Text = Data14.Recordset.Fields(2)
End If

'Text4.Text = K1
'Text5.Text = K2
'DTPicker1.Value = K1
'DTPicker2.Value = K2

Data1.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CKGL.mdb"
Data1.RecordSource = "select GYS.MC from GYS  GROUP BY GYS.MC"
Data1.Refresh

Data5.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data5.RecordSource = "select CWZY.摘要 from CWZY  GROUP BY CWZY.摘要"
Data5.Refresh

Data6.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data6.RecordSource = "select CWMC.科目名称 from CWMC WHERE CWMC.科目方向='借' GROUP BY CWMC.科目名称"
Data6.Refresh

Data7.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data7.RecordSource = "select CWMC.科目名称 from CWMC WHERE CWMC.科目方向='贷' GROUP BY CWMC.科目名称"
Data7.Refresh

Data8.DatabaseName = "d:\数据库\bfrz\" + ljb + "\SCZYJHD.mdb"
Data8.RecordSource = "select 简称 from khzl  GROUP BY 简称"
Data8.Refresh

Data9.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data9.RecordSource = "select FHY.MC from FHY GROUP BY FHY.MC"
Data9.Refresh

Data10.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data10.Refresh



Data2.DatabaseName = "d:\数据库\bfrz\" + ljb + "\CW.mdb"
Data2.Refresh
Label7.Caption = ""
For i = 0 To 4
DBCombo1(i).Text = ""
DBCombo2(i).Text = ""
DBCombo3(i).Text = ""
DBCombo4(i).Text = ""
DBCombo5(i).Text = ""
DBCombo6(i).Text = ""
DBCombo7(i).Text = ""
Text5(i).Text = ""
Next
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""

'MSFlexGrid1.ColWidth(13) = 0

End Sub
Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
Text4.Text = DTPicker1.Value
End Sub


Private Sub Label10_Click(Index As Integer)
Select Case Index
       Case Index
       KMBL = Index
       KMMC = 4
Formw62.Show
End Select

End Sub


Private Sub Label8_Click(Index As Integer)
Select Case Index
       Case Index
       KMBL = Index
       KMMC = 2
Formw62.Show
End Select
End Sub


Private Sub MSFlexGrid2_dblClick()
On Error Resume Next
rs = MSFlexGrid2.Row
If Data2.Recordset.EOF Then Exit Sub
Data2.Recordset.MoveFirst
Data2.Recordset.Move rs - 1
 DBCombo1(0).Text = Data2.Recordset.Fields(0)
 DBCombo2(0).Text = Data2.Recordset.Fields(1)
 DBCombo3(0).Text = Data2.Recordset.Fields(2)
 DBCombo4(0).Text = Data2.Recordset.Fields(3)
 DBCombo5(0).Text = Data2.Recordset.Fields(4)
 DBCombo6(0).Text = Data2.Recordset.Fields(5)
 DTPicker3.Value = Data2.Recordset.Fields(7)
 Text5(0).Text = Data2.Recordset.Fields(8)
 DBCombo7(1).Text = Data2.Recordset.Fields(9)
 DBCombo7(2).Text = Data2.Recordset.Fields(10)
 DBCombo7(4).Text = Data2.Recordset.Fields(11)
End Sub

Private Sub Text2_Change()
If Combo2.Text = "转账凭证" Then
Data2.RecordSource = "SELECT * FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLZZPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If
'''''''''''''''''''''''''''''''''''''''''''''
If Combo2.Text = "付款凭证" Then
Data2.RecordSource = "SELECT * FROM CLFKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLFKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

If Combo2.Text = "收款凭证" Then
Data2.RecordSource = "SELECT * FROM CLSKPZ WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSKPZ.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

If Combo2.Text = "成本凭证" Then
Data2.RecordSource = "SELECT * FROM CLSCCB WHERE 日期 BETWEEN CDATE('" & K1 & "') AND CDATE('" & K2 & "') AND CLSCCB.凭证号='" & Text2.Text & "'"
Data2.Refresh
End If

End Sub

Private Sub PZDY(PZLB As String, DH As String) ''''无标题

        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军染整软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：

'Select Case Mid(DH, 1, 1)
'       Case "4"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYyf.xls")
'        Excelapp.Sheets(1).Activate
'       Case "2"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYxf.xls")
'        Excelapp.Sheets(1).Activate
'       Case "5"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYzz.xls")
'        Excelapp.Sheets(1).Activate
'       Case "3"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYys.xls")
'        Excelapp.Sheets(1).Activate
'       Case "1"
'        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDYxs.xls")
'        Excelapp.Sheets(1).Activate
'End Select

        Excelapp.Workbooks.Open ("e:\Excel\染整\宝隆\PZDY.xls")
        Excelapp.Sheets(1).Activate


Data2.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 6) = Data2.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(2, 5) = PZLB
        Excelapp.ActiveSheet.Cells(3, 10) = Trim(DH)
        Excelapp.ActiveSheet.Cells(16, 2) = Data2.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(16, 7) = Data2.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(16, 10) = Data2.Recordset.Fields(9)
i = 5
Do While Not Data2.Recordset.EOF

        Excelapp.ActiveSheet.Cells(i, 1) = Data2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(i, 3) = Data2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 5) = Data2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 8) = Val(Data2.Recordset.Fields(5))
        
        i = i + 1
        
        Excelapp.ActiveSheet.Cells(i, 1) = Data2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(i, 3) = Data2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 5) = Data2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 10) = Val(Data2.Recordset.Fields(5))
        
i = i + 1
Data2.Recordset.MoveNext
Loop

Excelapp.ActiveWindow.Zoom = 100


        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing

End Sub



