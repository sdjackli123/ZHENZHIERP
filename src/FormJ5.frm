VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formj5 
   BackColor       =   &H00C0E0FF&
   Caption         =   "强制排产"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   Icon            =   "FormJ5.frx":0000
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc10 
      Height          =   330
      Left            =   2400
      Top             =   10680
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "FormJ5.frx":440A
      Height          =   390
      Left            =   12240
      TabIndex        =   49
      Top             =   4440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "名称"
      Text            =   "DataCombo4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6960
      ItemData        =   "FormJ5.frx":4420
      Left            =   15480
      List            =   "FormJ5.frx":4422
      TabIndex        =   47
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "排产"
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Text            =   "Text6"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6960
      TabIndex        =   40
      Text            =   "Text6"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Text            =   "Text6"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5040
      TabIndex        =   37
      Text            =   "Text5"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3600
      TabIndex        =   36
      Text            =   "Text4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查询条件"
      Height          =   1335
      Left            =   11040
      TabIndex        =   28
      Top             =   360
      Width           =   2175
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "锅号"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "布类"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   33
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "色别"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "日期"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "重量"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "客户"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Text            =   "FormJ5.frx":4424
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   11040
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   7320
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   615
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "查询"
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "客户刷新"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "排完"
      Height          =   375
      Left            =   11040
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "续排"
      Height          =   375
      Left            =   11040
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "双染棉"
      Height          =   375
      Left            =   15360
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "排产类别"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "染色"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "印花"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "FormJ5.frx":442C
      Height          =   330
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "客户名称"
      Text            =   "DataCombo3"
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   11040
      TabIndex        =   1
      Top             =   9120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo2"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FormJ5.frx":4441
      Height          =   390
      Left            =   11040
      TabIndex        =   2
      Top             =   5640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   688
      _Version        =   393216
      ListField       =   "车台编号"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   4200
      Top             =   10680
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
      Height          =   330
      Left            =   4560
      Top             =   10560
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
      Height          =   375
      Left            =   3360
      Top             =   10560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Height          =   375
      Left            =   4080
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Height          =   375
      Left            =   3600
      Top             =   10560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   375
      Left            =   3480
      Top             =   10440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Left            =   3960
      Top             =   10680
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
      Height          =   375
      Left            =   4440
      Top             =   10680
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
      Height          =   375
      Left            =   4320
      Top             =   10440
      Visible         =   0   'False
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "FormJ5.frx":4456
      Height          =   7815
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   10455
      _cx             =   18441
      _cy             =   13785
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   329711619
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   1118719
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   329711619
      CurrentDate     =   36892
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "染色方式"
      Height          =   375
      Index           =   5
      Left            =   12240
      TabIndex        =   48
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "排产锅数："
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
      Left            =   15480
      TabIndex        =   46
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   43
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "品名"
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   41
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "色别"
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   39
      Top             =   1680
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   5040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "重量范围"
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   35
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "请选择车台编号："
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
      Index           =   4
      Left            =   11040
      TabIndex        =   27
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "排产内容："
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
      Left            =   11040
      TabIndex        =   26
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "LABEL2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   25
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "如编号，请输入。"
      Height          =   255
      Left            =   11040
      TabIndex        =   24
      Top             =   7080
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "输入客户"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   23
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "锅号"
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   22
      Top             =   8760
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "刷新"
      Height          =   375
      Left            =   11040
      TabIndex        =   19
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "工序设置"
      Height          =   375
      Left            =   13320
      TabIndex        =   18
      Top             =   9600
      Width           =   1575
   End
End
Attribute VB_Name = "Formj5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Public zd As Integer: Public sh As String

Private Sub Command1_Click()
'On Error Resume Next

If Trim(Text1.Text) = "" Then
MsgBox ("排产内容不能为空")
Exit Sub
End If

If Trim(DataCombo1.Text) = "" Then
MsgBox ("请选择车台！")
Exit Sub
End If

    
    If Option1.value = True Then
    jbl = "单染"
    End If
    If Option2.value = True Then
    jbl = "双染"
    End If
    If Option3.value = True Then
    jbl = "三染"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''检测工序设置
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' 连接到数据库
    
    Set param = g_Cmd.CreateParameter("gh", adChar, adParamInput, 40, Trim(DataCombo2.Text))
    g_Cmd.Parameters.Append param
    
    Set param = g_Cmd.CreateParameter("jc", adSingle, adParamOutput)
    g_Cmd.Parameters.Append param
    
    g_Cmd.CommandType = adCmdStoredProc     ' 表示cmd的类型为存储过程
    g_Cmd.CommandText = "gxszjc"       ' 表示调用哪个存储过程"
    g_Cmd.Execute           ' 执行存储过程
    g_Cmd.Cancel

If Val(g_Cmd.Parameters("jc").value) = 0 Then
If MsgBox("请先设置工序,不设置工序就继续排缸吗？", vbYesNo) = vbNo Then Exit Sub
End If


If Text2.Text = "" Then
Adodc2.Recordset.Fields(zd + 1) = Text1.Text + DataCombo4
Else
If zd + 1 < Val(Text2.Text) Then
MsgBox ("wrong")
Exit Sub
Else
L = Val(Text2.Text)
i = zd + 1
Do Until i = L
Adodc2.Recordset.Fields(i) = Adodc2.Recordset.Fields(i - 1)
i = i - 1
Loop
Adodc2.Recordset.Fields(L) = Text1.Text + DataCombo4
End If
End If
Adodc2.Recordset.Update

If Option4.value = True Then
If Option1.value = True Then
sql1 = "UPDATE KPD SET ye=convert(nvarchar(120) ,getdate(),120),车台='" & DataCombo1.Text & "',zt='已染缸计划',ztbh='1001',scbh='1001',gz=convert(nvarchar(120) ,getdate(),120) WHERE 锅号='" & DataCombo2.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Text1.Text = ""
DataCombo2.Text = ""
End If
End If

If Option5.value = True Then
sql1 = "UPDATE KPD SET ye=convert(nvarchar(120),getdate(),120),车台='" & DataCombo1.Text & "',zt='已印花计划',ztbh='1001'scbh='1002',gz=convert(nvarchar(120) ,getdate(),120) WHERE 锅号='" & DataCombo2.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Text1.Text = ""
DataCombo2.Text = ""
End If

Call DataCombo1_Change
Call Command5_Click
Call gssx
MsgBox ("排产成功！")

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub pcxx(xx As String)
On Error Resume Next
    Adodc3.RecordSource = "select max(重量) from kpd where 锅号='" & xx & "'"
    Adodc3.Refresh
    gk = Adodc3.Recordset.Fields(0)
   
    '查找数量最多的锅号
    Adodc3.RecordSource = "select * from kpd where 锅号='" & xx & "' and 重量='" & gk & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.EOF Then Text1.Text = ""
    '同锅号合计重量
    Adodc5.RecordSource = "select sum(重量) as op from kpd where 锅号='" & xx & "'"
    Adodc5.Refresh
    If Adodc5.Recordset.EOF Then Text1.Text = ""
    '付值数组变量
    
    Text1.Text = Trim(Adodc3.Recordset.Fields(0)) + ":" + Trim(Adodc3.Recordset.Fields(3)) + ".." + Trim(Adodc3.Recordset.Fields(8)) + Trim(Adodc3.Recordset.Fields("色名")) + "\" + Trim(Format(Adodc5.Recordset.Fields(0), "#0.0")) + "kg" + "锅号" + Trim(Adodc3.Recordset.Fields(2))
   '添加标示
   '''' MsgBox (Adodc4.Recordset.Fields(0))
   sh = Adodc3.Recordset.Fields(8)
End Sub

Private Sub Command3_Click()
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
t1 = Format(Trim(DTPicker3.value), "yyyy-MM-dd")
t2 = Format(Trim(DTPicker4.value), "yyyy-MM-dd")
Adodc7.RecordSource = "select 客户名称  from kpd where ye='N' AND CONVERT(varchar(120),日期,23) between '" & t1 & "' and '" & t2 & "'  group by 客户名称"
Adodc7.Refresh
End Sub


Private Sub Command4_Click()
Formj31.DataCombo1 = DataCombo1
Formj31.Show
End Sub

Private Sub Command5_Click()
'On Error Resume Next
sql1 = ""

If Check2(1).value = 1 Then
sql1 = sql1 + "客户名称 like '%'+'" & DataCombo3.Text & "'+'%' and "
End If


If Check2(3).value = 1 Then
sql1 = sql1 + "合计重量 between '" & Text4 & "' and '" & Text5 & "' and "
End If

If Check2(4).value = 1 Then
sql1 = sql1 + "日期 between cast('" & DTPicker3.value & "' as datetime) and cast('" & DTPicker4.value & "' as datetime) and "
End If

If Check2(7).value = 1 Then
sql1 = sql1 + "品名 like '%'+'" & Text3.Text & "'+'%' and "
End If

If Check2(6).value = 1 Then
sql1 = sql1 + "锅号 like '%'+'" & Text7.Text & "'+'%' and "
End If

If Check2(5).value = 1 Then
sql1 = sql1 + "色别 like '%'+'" & Text6.Text & "'+'%' and "
End If


If sql1 = "" Then
MsgBox ("请选择查询条件")
Exit Sub
End If
sql1 = Left$(Trim(sql1), Len(Trim(sql1)) - 4)


If Option4.value = True Then
Adodc8.RecordSource = "select 客户名称,锅号,色别,品名,匹数,重量,类别,染色要求,日期,合计重量  from v_jhkpd where (" + sql1 + ")  order by 日期,色别,锅号"
Adodc8.Refresh
Adodc4.RecordSource = "select 车台编号  from jhb where (" + sql1 + ")  order by 车台编号"
'Adodc4.Refresh
End If

If Option5.value = True Then
Adodc8.RecordSource = "select 客户名称,锅号,色别,品名,匹数,重量,类别,染色要求,日期,合计重量  from v_jhkpd where (" + sql1 + ")  order by 日期,色别,锅号"
Adodc8.Refresh
End If

VSFlexGrid1.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1, False, 30

VSFlexGrid1.ColWidth(0) = 200
VSFlexGrid1.ColWidth(9) = 1200

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 800
Next
End If

Call gssx
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()
On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Exit Sub
Else
zd = 0
For i = 1 To 30
If Adodc2.Recordset.Fields(i) <> "" Then
zd = zd + 1
End If
Next
Label2.Caption = DataCombo1.Text + Space(4) + "共排产" + Space(3) + Str(zd) + "锅"
End If

End Sub

Private Sub DataCombo1_Click(Area As Integer)
On Error Resume Next
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select * from jhb where 车台编号='" & DataCombo1.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Exit Sub
Else
zd = 0
For i = 1 To 30
If Adodc2.Recordset.Fields(i) <> "" Then
zd = zd + 1
End If
Next
Label2.Caption = DataCombo1.Text + Space(4) + "共排产" + Space(3) + Str(zd) + "锅"
End If

End Sub

Private Sub Form_Load()

'On Error Resume Next
DTPicker3.value = Date
DTPicker4.value = Date
Option4.value = True


Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 车台编号  from ct group by 车台编号"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "jhb"
Adodc2.Refresh

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select 车台编号  from ct group by 车台编号"
Adodc6.Refresh

Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select 客户名称  from kpd where ye='N' and 锅号<>''   group by 客户名称"
Adodc7.Refresh

Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select distinct 名称,序号  from cjrsfs order by 序号"
Adodc10.Refresh

Option1.value = True


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label2.Caption = ""

DataCombo1.Text = ""
DataCombo2.Text = ""
DataCombo3.Text = ""
DataCombo4.Text = ""

VSFlexGrid1.ColWidth(0) = 200
For i = 1 To 9
VSFlexGrid1.ColWidth(i) = 1600
Next

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 800
Next
End If

End Sub


Private Sub Label8_Click()
If Option4.value = True Then
Adodc8.RecordSource = "select 客户名称,锅号,色别,品名,匹数,重量,类别,备注,日期  from kpd where ye='N' and 锅号='" & DataCombo2.Text & "' and 锅号 not like 'Y'  order by 客户名称,日期,锅号"
Adodc8.Refresh
End If
If Option5.value = True Then
Adodc8.RecordSource = "select 客户名称,锅号,色别,品名,匹数,重量,类别,备注,日期  from kpd where ye='N' and 锅号='" & DataCombo2.Text & "' and 锅号 like 'Y' AND 锅号 not like 'D' AND 锅号 not like 'F'  order by 客户名称,日期,锅号"
Adodc8.Refresh
End If
Call pcxx(DataCombo2.Text)
End Sub

Private Sub Label9_Click()
Formd332.Text1 = DataCombo2
Formd332.Show
End Sub

Private Sub gssx()

Adodc9.RecordSource = "select 车台编号 from jhb order by 车台编号"
Adodc9.Refresh

List2.Clear

If Not Adodc9.Recordset.EOF Then
Adodc9.Recordset.MoveFirst
Do While Not Adodc9.Recordset.EOF
Adodc4.RecordSource = "select * from jhb where 车台编号='" & Adodc9.Recordset.Fields(0) & "'"
Adodc4.Refresh
If Not Adodc4.Recordset.EOF Then
zd = 0
For i = 1 To 30
If Adodc4.Recordset.Fields(i) <> "" Then
zd = zd + 1
End If
Next
List2.AddItem Adodc9.Recordset.Fields(0) + Space(2) + "--" + Space(2) + Trim(zd) + "锅"
End If
Adodc9.Recordset.MoveNext
Loop
End If
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc8.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc8.Recordset.MoveFirst
Adodc8.Recordset.Move rs - 1
DataCombo2.Text = Adodc8.Recordset.Fields(1)
If IsNull(Adodc8.Recordset.Fields(1)) Then Exit Sub
Call pcxx(Adodc8.Recordset.Fields(1))
End Sub


