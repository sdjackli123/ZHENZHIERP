VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "ë�����"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Forma2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   20520
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   480
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc17 
      Height          =   375
      Left            =   8880
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Adodc17"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   52
      Text            =   "Forma2.frx":440A
      Top             =   240
      Width           =   11055
   End
   Begin MSAdodcLib.Adodc Adodc16 
      Height          =   375
      Left            =   3480
      Top             =   9960
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "Adodc16"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ѯ"
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
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc14 
      Height          =   330
      Left            =   6480
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
      Caption         =   "Adodc14"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   6840
      Top             =   10080
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
      Caption         =   "Adodc13"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ǩ��ӡ"
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
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "�ͻ���Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Index           =   1
      Left            =   4200
      TabIndex        =   22
      Top             =   1800
      Width           =   14655
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   360
         Left            =   1440
         TabIndex        =   54
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo2"
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1560
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":4410
         Height          =   360
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "���"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":4425
         Height          =   360
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "pm"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   3
         Left            =   5640
         TabIndex        =   5
         Top             =   2520
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   4
         Left            =   5640
         TabIndex        =   6
         Top             =   3120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   5
         Left            =   9000
         TabIndex        =   7
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   6
         Left            =   9000
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   7
         Left            =   11760
         TabIndex        =   10
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   8
         Left            =   11760
         TabIndex        =   11
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":443A
         Height          =   360
         Index           =   9
         Left            =   9000
         TabIndex        =   8
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "���λ��"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":444F
         Height          =   360
         Index           =   10
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "xm"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   11
         Left            =   1440
         TabIndex        =   4
         Top             =   3120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":4464
         Height          =   360
         Index           =   12
         Left            =   11760
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "ny"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":4479
         Height          =   360
         Index           =   16
         Left            =   11760
         TabIndex        =   49
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "����"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":448F
         Height          =   360
         Index           =   17
         Left            =   5640
         TabIndex        =   58
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "����"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma2.frx":44A5
         Height          =   360
         Index           =   18
         Left            =   5640
         TabIndex        =   59
         Top             =   2040
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "����"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   19
         Left            =   5640
         TabIndex        =   62
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   20
         Left            =   5640
         TabIndex        =   63
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ƥ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   4440
         TabIndex        =   61
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����ƥ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4440
         TabIndex        =   60
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4440
         TabIndex        =   57
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4440
         TabIndex        =   56
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ɫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "˾��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   10920
         TabIndex        =   48
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ë������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "�ͻ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ͬ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ë������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4440
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   10920
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7800
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ë������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ë��ƥ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4440
         TabIndex        =   28
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���λ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "���ϵ�λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   10920
         TabIndex        =   25
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�޸ı���"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1335
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
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
      Left            =   16920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��һ����"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ݴ�ӡ"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��������"
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
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3720
      Top             =   10320
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Forma2.frx":44BB
      Height          =   3975
      Left            =   4320
      TabIndex        =   36
      Top             =   6720
      Width           =   18735
      _cx             =   33046
      _cy             =   7011
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4200
      Top             =   10200
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4080
      Top             =   10200
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4680
      Top             =   10080
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4680
      Top             =   10080
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4320
      Top             =   10320
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4200
      Top             =   10320
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   4800
      Top             =   10200
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Top             =   10200
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   3960
      Top             =   10080
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   13
      Left            =   1440
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   14
      Left            =   5280
      TabIndex        =   38
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   360
      Index           =   15
      Left            =   5040
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Text            =   "DataCombo1"
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Forma2.frx":44D1
      Height          =   8415
      Left            =   240
      TabIndex        =   40
      Top             =   1080
      Width           =   3615
      _cx             =   6376
      _cy             =   14843
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
   Begin MSAdodcLib.Adodc Adodc11 
      Height          =   330
      Left            =   2400
      Top             =   10200
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      Left            =   840
      Top             =   10080
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma2.frx":44E7
      Height          =   855
      Left            =   15360
      TabIndex        =   47
      Top             =   7680
      Width           =   2775
      _cx             =   4895
      _cy             =   1508
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
      FormatString    =   $"Forma2.frx":44FD
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
   Begin MSAdodcLib.Adodc Adodc15 
      Height          =   330
      Left            =   4920
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
      Caption         =   "Adodc14"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "������ϸ"
      Height          =   1575
      Left            =   8040
      TabIndex        =   51
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "�뾡��ѡ��Ʒ��"
      Height          =   375
      Left            =   240
      TabIndex        =   50
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��ǰ���ݺ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   43
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   960
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����Ʒ��"
      Height          =   375
      Left            =   3960
      TabIndex        =   41
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Forma2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim conn As ADODB.Connection
Dim RD As ADODB.Recordset
Dim cdbhf As Integer

Private Sub Command1_Click()
Forma1111.Show
End Sub

Private Sub Command12_Click()
On Error Resume Next
If DataCombo1(14).Text = "" Then Exit Sub

Adodc14.RecordSource = "select isnull(count(ip),0) from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc14.Refresh
If Not Adodc14.Recordset.EOF Then

Call mprk(Adodc8, DataCombo1(14).Text)


End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call mprkbqdy(Adodc8, DataCombo1(14), DataCombo1(7))
End Sub

Private Sub Command33_Click()
Unload Me
End Sub

Private Sub Command5_Click()
wwdm = 5
Forma6.Show
End Sub

Private Sub Command6_Click()
    Dim SumCM As Double
    Dim CountCM As Integer
    Dim SumLing As Double
    Dim CountLing As Integer
    Dim inputText As String
    Dim lines As Variant
    Dim i As Integer
    Dim isLing As Boolean

    ' ��ȡ�����ı�
    inputText = Text3.Text
    ' ʹ�ûس�����vbCrLf������ı�Ϊ��
    lines = Split(inputText, vbCrLf)

    ' ѭ������ÿ���ı�
    For i = LBound(lines) To UBound(lines)
        ' ����Ƿ��� �� ��
        If InStr(lines(i), "��") > 0 Then
            isLing = True
        Else
            isLing = False
        End If

        ' ��ÿ���ı�����ȡ���ֲ���
        Dim numbers As Variant
        numbers = Split(lines(i), " ")

        ' ѭ������ÿ������
        Dim j As Integer
        For j = LBound(numbers) To UBound(numbers)
            ' ���Խ��ַ���ת��Ϊ Double
            Dim Number As Double
            If IsNumeric(numbers(j)) Then
                Number = CDbl(numbers(j))
                If isLing Then
                    ' �ۼ�����ܺͺ͸���
                    SumLing = SumLing + Number
                    CountLing = CountLing + 1
                Else
                    ' �ۼ�cm���ܺͺ͸���
                    SumCM = SumCM + Number
                    CountCM = CountCM + 1
                End If
            End If
        Next j
    Next i

    ' ����cm���ۼӵĽ����ֵ�� DataCombo1(17)
    DataCombo1(17).Text = SumCM
    ' �����ָ�����ֵ�� DataCombo1(20)
    DataCombo1(20).Text = CountCM
    ' �� �� ���ۼӵĽ����ֵ�� DataCombo1(18)
    DataCombo1(18).Text = SumLing
    ' �����ָ�����ֵ�� DataCombo1(19)
    DataCombo1(19).Text = CountLing
    ' �� SumCM + SumLing ������λС����ֵ�� DataCombo1(3)
    DataCombo1(3).Text = Format(SumCM + SumLing, "0.00")
    ' �� CountCM + CountLing ��ֵ�� DataCombo1(4)
    DataCombo1(4).Text = CountCM + CountLing
End Sub






Private Sub Command7_Click()
On Error Resume Next
If InStr(DataCombo1(14), yhdm) > 0 Then
Adodc5.RecordSource = "select   * from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc5.Refresh
Adodc15.RecordSource = "select   dh as ����,�ͻ�����,����,ë�߷���,���� as ë������,ë��ƥ��,ë������,��Լ��,��ע,����,���ݺ�,���λ��,ny as ���ϵ�λ,������,ҵ��,������,��������,����ƥ��,��ƥ�� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc15.Refresh
Adodc13.RecordSource = "select  sum(ë��ƥ��) as �ϼ�ƥ��,sum(ë������) as �ϼ����� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc13.Refresh
Adodc10.RecordSource = "SELECT ip FROM CKGL WHERE ���ݺ�='" & DataCombo1(14).Text & "' order by ip desc"
Adodc10.Refresh
If Adodc10.Recordset.EOF Then
DataCombo1(7).Text = 1
Else
DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
End If
DataCombo1(8).Text = Date
End If
Call gjsx
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub


Private Sub Command11_Click()
 Call Command6_Click
'On Error Resume Next
'If DataCombo1(10).Text = "" Then
'MsgBox ("��ѡ�����ˣ�")
'Exit Sub
'End If
Dim i As Integer

If DataCombo1(3).Text = "" Or DataCombo1(4).Text = "" Then
MsgBox ("������ƥ������Ϊ�գ���")
Exit Sub
End If

'If Mid(DataCombo1(14), 1, 1) <> yhdm Then
'MsgBox ("���ݺŴ������û����벻����")
'Exit Sub
'End If

'If Len(DataCombo1(14)) <> 8 Then
'MsgBox ("���ݺŲ�����8λ����")
'Exit Sub
'End If

If DataCombo1(8).Text = "" Then
MsgBox ("���������ڣ�")
Exit Sub
End If


    Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "mpckgl('" & DataCombo1(0).Text & "','" & DataCombo1(1).Text & "','" & DataCombo1(2).Text & "','" & DataCombo1(3).Text & "','" & DataCombo1(4).Text & "','" & DataCombo1(5).Text & "','" & DataCombo1(6).Text & "','" & DataCombo1(7).Text & "','" & DataCombo1(8).Text & "','" & DataCombo1(9).Text & "','" & DataCombo1(10).Text & "','" & DataCombo1(11).Text & "','" & DataCombo1(12).Text & "','" & DataCombo1(13).Text & "','" & DataCombo1(14).Text & "','" & Text3.Text & " ','" & DataCombo1(16).Text & "','" & DataCombo2.Text & "','" & DataCombo1(17).Text & "','" & DataCombo1(18).Text & "','" & DataCombo1(19).Text & "','" & DataCombo1(20).Text & "')"      ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel


Adodc10.Refresh
DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
DataCombo1(8).Text = Date


Adodc3.Refresh
Adodc5.Refresh
Adodc15.RecordSource = "select   dh as ����,������ϸ,�ͻ�����,����,ë�߷���,���� as ë������,ë��ƥ��,ë������,��Լ��,��ע,����,���ݺ�,���λ��,ny as ���ϵ�λ,������,ҵ��,��ɫ,������,��������,����ƥ��,��ƥ�� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc15.Refresh
Adodc13.RecordSource = "select  sum(ë��ƥ��) as �ϼ�ƥ��,sum(ë������) as �ϼ����� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc13.Refresh

Call gjsx


DataCombo1(0).SetFocus

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Integer

'If Mid(DataCombo1(14), 1, 1) <> yhdm Then
'MsgBox ("���ݺŴ������û����벻����")
'Exit Sub
'End If

If Len(DataCombo1(14)) <> 8 Then
MsgBox ("���ݺŲ�����8λ����")
Exit Sub
End If

If DataCombo1(3).Text = "" Or DataCombo1(4).Text = "" Then
MsgBox ("������ƥ������Ϊ�գ���")
Exit Sub
End If

If DataCombo1(8).Text = "" Then
MsgBox ("���������ڣ�")
Exit Sub
End If

If DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1 Then
MsgBox ("������IP��")
Exit Sub
End If

For i = 0 To 14
Adodc5.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc5.Recordset.Fields(15) = Text3.Text
Adodc5.Recordset.Fields(16) = DataCombo1(16).Text
Adodc5.Recordset.Fields(18) = DataCombo1(17).Text
Adodc5.Recordset.Fields(19) = DataCombo1(18).Text
Adodc5.Recordset.Fields(20) = DataCombo1(19).Text
Adodc5.Recordset.Fields(21) = DataCombo1(20).Text
Adodc5.Recordset.Fields(17) = DataCombo2.Text
Adodc5.Recordset.Update
Adodc5.Refresh
Adodc15.RecordSource = "select   dh as ����,������ϸ,�ͻ�����,����,ë�߷���,���� as ë������,ë��ƥ��,ë������,��Լ��,��ע,����,���ݺ�,���λ��,ny as ���ϵ�λ,������,ҵ��,������,��������,����ƥ��,��ƥ�� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc15.Refresh
Adodc13.RecordSource = "select  sum(ë��ƥ��) as �ϼ�ƥ��,sum(ë������) as �ϼ����� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc13.Refresh

Adodc10.Refresh
If Adodc10.Recordset.EOF Then
DataCombo1(7).Text = 1
Else
DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
End If
DataCombo1(8).Text = Date

Call gjsx


DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
DataCombo1(8).Text = Date
DataCombo1(0).SetFocus
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim i As Integer
If MsgBox("ȷ��ɾ����", vbYesNo) = vbNo Then Exit Sub
Adodc5.Recordset.Delete

Adodc13.RecordSource = "select  sum(ë��ƥ��) as �ϼ�ƥ��,sum(ë������) as �ϼ����� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc13.Refresh
Adodc15.RecordSource = "select   dh as ����,�ͻ�����,����,ë�߷���,���� as ë������,ë��ƥ��,ë������,��Լ��,��ע,����,���ݺ�,���λ��,ny as ���ϵ�λ,������,ҵ��,������,��������,����ƥ��,��ƥ�� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc15.Refresh
Adodc5.Refresh

Adodc10.Refresh
If Adodc10.Recordset.EOF Then
DataCombo1(7).Text = 1
Else
DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
End If
DataCombo1(8).Text = Date
Call gjsx
DataCombo1(0).SetFocus
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub


Private Sub Command8_Click()
'On Error Resume Next  ' ������ִ��������ִ����һ��

'On Error Resume Next  ' ������ִ��������ִ����һ��

Dim yearMonth As String
yearMonth = Format(Date, "yyMM")  ' ��ȡ��ǰ���ڲ���ʽ��Ϊ "yyMM" ��ʽ

' �������ݿ������ַ���
Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
' ���ò�ѯ��䣬��ȡ��ǰ���¶�Ӧ��8λ���ݺŵ�������ֲ���
Adodc17.RecordSource = "SELECT ISNULL(MAX(CAST(SUBSTRING(���ݺ�, 5, 4) AS INT)), 0) AS h FROM ckgl WHERE SUBSTRING(���ݺ�, 1, 4)='" & yearMonth & "' AND LEN(���ݺ�) = 8"
Adodc17.Refresh  ' ˢ�����ݿؼ���ִ�в�ѯ

Dim newBillNumber As String  ' �����µĵ��ݺű���

If Adodc17.Recordset.EOF Then  ' �����¼��Ϊ��
    newBillNumber = yearMonth & "0001"  ' �µĵ��ݺ�Ϊ��ǰ���¼��� "0001"
Else
    Dim nextNumber As Long  ' ������һ�����ݺ����ֲ��ֵı�������ʹ��Long�����Ա������
    nextNumber = CLng(Adodc17.Recordset.Fields(0)) + 1  ' ��ȡ��ǰ������ֲ��ֲ���1
    newBillNumber = yearMonth & Format(nextNumber, "0000")  ' ��ʽ���µĵ��ݺ����ֲ���Ϊ4λ����
End If


    DataCombo1(14).Text = newBillNumber  ' ���µĵ��ݺŸ�ֵ���ؼ�

    For i = 0 To 7  ' �����һ�������������
        DataCombo1(i).Text = ""
    Next
    For i = 9 To 13  ' �����һ�������������
        DataCombo1(i).Text = ""
    Next
    DataCombo1(7).Text = 1
    DataCombo1(16).Text = ""
    DataCombo1(17).Text = ""
    DataCombo1(18).Text = ""
    DataCombo1(19).Text = ""
    DataCombo1(20).Text = ""
    Text3.Text = ""
    Call Command7_Click
    
End Sub


Private Sub DataCombo1_Change(Index As Integer)
On Error Resume Next
Dim i As Integer
Select Case Index

Case 14
'If InStr(DataCombo1(14), yhdm) > 0 Or yhm = "root" Then
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select   * from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc5.Refresh
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.RecordSource = "select  dh as ����,������ϸ,�ͻ�����,����,ë�߷���,���� as ë������,ë��ƥ��,ë������,��Լ��,��ע,����,���ݺ�,���λ��,ny as ���ϵ�λ,������,ҵ��,������,��������,����ƥ��,��ƥ�� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc15.Refresh
Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT ip FROM CKGL WHERE ���ݺ�='" & DataCombo1(14).Text & "' order by ip desc"
Adodc10.Refresh
DataCombo1(7).Text = Val(Adodc10.Recordset.Fields(0)) + 1
DataCombo1(8).Text = Date
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "select  sum(ë��ƥ��) as �ϼ�ƥ��,sum(ë������) as �ϼ����� from ckgl where ���ݺ�='" & DataCombo1(14).Text & "' "
Adodc13.Refresh
'End If
End Select

Call gjsx
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Dim i As Integer
Select Case Index
Case 0
t1 = Format(Date - 180, "yyyy-mm-dd")
t2 = Format(Date, "yyyy-mm-dd")
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct ���� from ckgl where �ͻ�����='" & DataCombo1(0).Text & "' order by ���� DESC"
Adodc11.Refresh
Case 10
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT xm ,IP from FZR"
Adodc4.Refresh
End Select

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 300
Next
End If


Call gjsx
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub Form_Load()
On Error Resume Next
kg = 1
Text1.Text = 0
kkf = 0
DH = 1
Dim i As Integer


cdbhf = cdbh
For i = 0 To 20
DataCombo1(i).Text = ""
Next
DataCombo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
DataCombo1(8).Text = Date


Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

DataCombo1(10).Text = ""                      ''''���ø�����=�û�����
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "SELECT xm as ����������,IP from FZR"
Adodc4.Refresh
If Not Adodc14.Recordset.EOF Then
DataCombo1(10) = Adodc4.Recordset.Fields(0)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from KHZL  where ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc1.Refresh

Frame3.Visible = False

Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select ���λ�� from ckgl group by ���λ��"
Adodc3.Refresh

Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.RecordSource = "select distinct ���� from ywf"
Adodc16.Refresh

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"


Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select ny from ckgl   group by ny"  ''�������ϵ�λ
Adodc6.Refresh


 'On Error Resume Next  ' ������ִ��������ִ����һ��

Dim yearMonth As String
yearMonth = Format(Date, "yyMM")  ' ��ȡ��ǰ���ڲ���ʽ��Ϊ "yyMM" ��ʽ

' �������ݿ������ַ���
Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
' ���ò�ѯ��䣬��ȡ��ǰ���¶�Ӧ��8λ���ݺŵ�������ֲ���
Adodc17.RecordSource = "SELECT ISNULL(MAX(CAST(SUBSTRING(���ݺ�, 5, 4) AS INT)), 0) AS h FROM ckgl WHERE SUBSTRING(���ݺ�, 1, 4)='" & yearMonth & "' AND LEN(���ݺ�) = 8"
Adodc17.Refresh  ' ˢ�����ݿؼ���ִ�в�ѯ

Dim newBillNumber As String  ' �����µĵ��ݺű���

If Adodc17.Recordset.EOF Then  ' �����¼��Ϊ��
    newBillNumber = yearMonth & "0001"  ' �µĵ��ݺ�Ϊ��ǰ���¼��� "0001"
Else
    Dim nextNumber As Long  ' ������һ�����ݺ����ֲ��ֵı�������ʹ��Long�����Ա������
    nextNumber = CLng(Adodc17.Recordset.Fields(0)) + 1  ' ��ȡ��ǰ������ֲ��ֲ���1
    newBillNumber = yearMonth & Format(nextNumber, "0000")  ' ��ʽ���µĵ��ݺ����ֲ���Ϊ4λ����
End If

    DataCombo1(14).Text = newBillNumber  ' ���µĵ��ݺŸ�ֵ���ؼ�

Call Command8_Click
Call Label7_Click

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "SELECT MAX(ip) FROM CKGL WHERE ���ݺ�='" & DataCombo1(14).Text & "'"
Adodc10.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(7) = 1500
VSFlexGrid1.ColWidth(9) = 2000


If Adodc5.Recordset.EOF Then
VSFlexGrid1.TextMatrix(0, 0) = "��¼��"
GoTo 100
End If
Adodc5.Recordset.MoveLast
VSFlexGrid1.TextMatrix(0, 0) = "��¼��"
For i = 1 To Adodc5.Recordset.RecordCount
VSFlexGrid1.TextMatrix(i, 0) = i
Next
100:
DataCombo1(6).Text = ""

Adodc10.Refresh
DataCombo1(7).Text = 1
DataCombo1(8).Text = Date
DataCombo1(13).TabIndex = 0
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

If Len(yhdm) <> 1 Then
MsgBox ("����˻������ʽ����������")
Command7.Enabled = False
Command8.Enabled = False
Command11.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
DataCombo1(14).Enabled = True
End If

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

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case 3
beizhu = 52
Forma113.Show
End Select
End Sub

Private Sub Label5_Click()
Formy123.Text1 = DataCombo1(13)
Formy123.Show
End Sub

Private Sub Label7_Click()
DataCombo1(8).Text = Date
End Sub

Private Sub Text1_Change()
On Error Resume Next
Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select pm from pm  where pm like '%'+'" & Text1.Text & "'+'%' group by pm"
Adodc7.Refresh


t1 = Format(Date - 180, "yyyy-mm-dd")
t2 = Format(Date, "yyyy-mm-dd")
Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct ���� from ckgl where �ͻ�����='" & DataCombo1(0).Text & "' and ���� like '%'+'" & Text1 & "'+'%' and CONVERT(varchar,����, 23) between '" & t1 & "' and '" & t2 & "'  order by ���� DESC"
Adodc11.Refresh

If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 300
Next
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ��� from KHZL where ����  like '%'+'" & Text2 & "'+'%' and ip like '%'+'" & yhxx & "'+'%' group by ���"
Adodc1.Refresh
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc5.Recordset.EOF Then Exit Sub
Adodc5.Recordset.MoveFirst
Adodc5.Recordset.Move rs - 1
     For i = 0 To 14 ''''''''''''''''''''''���ݺŲ���
     'If i = 14 Then i = 15
     DataCombo1(i).Text = Adodc5.Recordset.Fields(i)
     Next
     Text3.Text = Adodc5.Recordset.Fields(15)
     DataCombo1(16).Text = Adodc5.Recordset.Fields(16)
      DataCombo1(17).Text = Adodc5.Recordset.Fields(18)
       DataCombo1(18).Text = Adodc5.Recordset.Fields(19)
        DataCombo1(19).Text = Adodc5.Recordset.Fields(20)
       DataCombo1(20).Text = Adodc5.Recordset.Fields(21)
       DataCombo2.Text = Adodc5.Recordset.Fields(17)
Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
rs = VSFlexGrid2.Row
If Adodc11.Recordset.EOF Then Exit Sub
Adodc11.Recordset.MoveFirst
Adodc11.Recordset.Move rs - 1
DataCombo1(1).Text = Adodc11.Recordset.Fields(0)
End Sub

Private Sub gjsx()
On Error Resume Next
VSFlexGrid1.ColWidth(0) = 200
For i = 1 To 15
VSFlexGrid1.ColWidth(i) = 1100
Next
If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 400
Next
End If

VSFlexGrid2.ColWidth(0) = 200
VSFlexGrid2.ColWidth(1) = 6000
If VSFlexGrid2.Rows > 1 Then
For i = 1 To VSFlexGrid2.Rows - 1
VSFlexGrid2.RowHeight(i) = 400
Next
End If

End Sub


