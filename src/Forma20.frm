VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma20 
   BackColor       =   &H00C0E0FF&
   Caption         =   "织布计划"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Forma20.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   19076
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "加工信息"
      TabPicture(0)   =   "Forma20.frx":440A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(26)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(25)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(24)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(23)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(22)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label3(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(21)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(20)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(17)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(19)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(18)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label2(16)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2(15)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label3(7)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label2(14)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label2(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label2(12)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label3(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label2(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label2(10)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label2(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label3(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label2(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label2(2)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label2(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label2(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label2(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label2(8)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label2(6)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label3(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label3(2)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label3(3)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label3(4)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label3(6)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label9"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label3(16)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label2(29)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label3(19)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "DataCombo1(39)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "DataCombo1(38)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "DataCombo1(37)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "DataCombo1(36)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "DataCombo1(35)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "DataCombo1(34)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "DataCombo1(33)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "DataCombo1(32)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "DataCombo1(31)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "DataCombo1(29)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "DTPicker2"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "DTPicker1"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "DataCombo1(28)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "DataCombo1(27)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "DataCombo1(26)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "DataCombo1(25)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "DataCombo1(24)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "DataCombo1(23)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "DataCombo1(22)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "DataCombo1(21)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "DataCombo1(20)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "DataCombo1(19)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "DataCombo1(18)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "DataCombo1(17)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "DataCombo1(16)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "DataCombo1(15)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "DataCombo1(14)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "DataCombo1(13)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "DataCombo1(12)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "DataCombo1(11)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "DataCombo1(10)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "DataCombo1(9)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "DataCombo1(8)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "DataCombo1(7)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "DataCombo1(6)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "DataCombo1(5)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "DataCombo1(4)"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "DataCombo1(3)"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "DataCombo1(2)"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "DataCombo1(1)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Adodc10"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Adodc9"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "Adodc8"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "Adodc7"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).Control(90)=   "Adodc6"
      Tab(0).Control(90).Enabled=   0   'False
      Tab(0).Control(91)=   "Adodc5"
      Tab(0).Control(91).Enabled=   0   'False
      Tab(0).Control(92)=   "Adodc4"
      Tab(0).Control(92).Enabled=   0   'False
      Tab(0).Control(93)=   "Adodc3"
      Tab(0).Control(93).Enabled=   0   'False
      Tab(0).Control(94)=   "Adodc2"
      Tab(0).Control(94).Enabled=   0   'False
      Tab(0).Control(95)=   "VSFlexGrid1"
      Tab(0).Control(95).Enabled=   0   'False
      Tab(0).Control(96)=   "Adodc1"
      Tab(0).Control(96).Enabled=   0   'False
      Tab(0).Control(97)=   "DataCombo1(0)"
      Tab(0).Control(97).Enabled=   0   'False
      Tab(0).Control(98)=   "Adodc11"
      Tab(0).Control(98).Enabled=   0   'False
      Tab(0).Control(99)=   "Adodc12"
      Tab(0).Control(99).Enabled=   0   'False
      Tab(0).Control(100)=   "Text4"
      Tab(0).Control(100).Enabled=   0   'False
      Tab(0).Control(101)=   "Command16"
      Tab(0).Control(101).Enabled=   0   'False
      Tab(0).Control(102)=   "Combo1"
      Tab(0).Control(102).Enabled=   0   'False
      Tab(0).Control(103)=   "Command15"
      Tab(0).Control(103).Enabled=   0   'False
      Tab(0).Control(104)=   "Command9"
      Tab(0).Control(104).Enabled=   0   'False
      Tab(0).Control(105)=   "Command6"
      Tab(0).Control(105).Enabled=   0   'False
      Tab(0).Control(106)=   "Command5"
      Tab(0).Control(106).Enabled=   0   'False
      Tab(0).Control(107)=   "Text3"
      Tab(0).Control(107).Enabled=   0   'False
      Tab(0).Control(108)=   "Text2"
      Tab(0).Control(108).Enabled=   0   'False
      Tab(0).Control(109)=   "Text1"
      Tab(0).Control(109).Enabled=   0   'False
      Tab(0).Control(110)=   "Command1"
      Tab(0).Control(110).Enabled=   0   'False
      Tab(0).Control(111)=   "Command12"
      Tab(0).Control(111).Enabled=   0   'False
      Tab(0).Control(112)=   "Command8"
      Tab(0).Control(112).Enabled=   0   'False
      Tab(0).Control(113)=   "Command7"
      Tab(0).Control(113).Enabled=   0   'False
      Tab(0).Control(114)=   "Command33"
      Tab(0).Control(114).Enabled=   0   'False
      Tab(0).Control(115)=   "Command11"
      Tab(0).Control(115).Enabled=   0   'False
      Tab(0).Control(116)=   "Command4"
      Tab(0).Control(116).Enabled=   0   'False
      Tab(0).Control(117)=   "Command2"
      Tab(0).Control(117).Enabled=   0   'False
      Tab(0).Control(118)=   "Adodc13"
      Tab(0).Control(118).Enabled=   0   'False
      Tab(0).Control(119)=   "Adodc14"
      Tab(0).Control(119).Enabled=   0   'False
      Tab(0).Control(120)=   "Adodc15"
      Tab(0).Control(120).Enabled=   0   'False
      Tab(0).Control(121)=   "Command3"
      Tab(0).Control(121).Enabled=   0   'False
      Tab(0).Control(122)=   "Text6"
      Tab(0).Control(122).Enabled=   0   'False
      Tab(0).Control(123)=   "Adodc16"
      Tab(0).Control(123).Enabled=   0   'False
      Tab(0).Control(124)=   "Adodc17"
      Tab(0).Control(124).Enabled=   0   'False
      Tab(0).Control(125)=   "Command10"
      Tab(0).Control(125).Enabled=   0   'False
      Tab(0).Control(126)=   "Text7"
      Tab(0).Control(126).Enabled=   0   'False
      Tab(0).Control(127)=   "Adodc18"
      Tab(0).Control(127).Enabled=   0   'False
      Tab(0).Control(128)=   "Text8"
      Tab(0).Control(128).Enabled=   0   'False
      Tab(0).ControlCount=   129
      TabCaption(1)   =   "价格信息"
      TabPicture(1)   =   "Forma20.frx":4426
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5(6)"
      Tab(1).Control(1)=   "Text5(5)"
      Tab(1).Control(2)=   "Text5(4)"
      Tab(1).Control(3)=   "Text5(3)"
      Tab(1).Control(4)=   "Text5(2)"
      Tab(1).Control(5)=   "Text5(1)"
      Tab(1).Control(6)=   "Text5(0)"
      Tab(1).Control(7)=   "Label3(18)"
      Tab(1).Control(8)=   "Label3(17)"
      Tab(1).Control(9)=   "Label3(15)"
      Tab(1).Control(10)=   "Label3(14)"
      Tab(1).Control(11)=   "Label2(28)"
      Tab(1).Control(12)=   "Label3(13)"
      Tab(1).Control(13)=   "Label2(27)"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "工艺信息"
      TabPicture(2)   =   "Forma20.frx":4442
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataCombo1(30)"
      Tab(2).Control(1)=   "Label2(7)"
      Tab(2).Control(2)=   "Image1"
      Tab(2).ControlCount=   3
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   7320
         TabIndex        =   127
         Text            =   "Text8"
         Top             =   840
         Width           =   2775
      End
      Begin MSAdodcLib.Adodc Adodc18 
         Height          =   330
         Left            =   1800
         Top             =   10320
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Adodc18"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   6
         Left            =   -66480
         TabIndex        =   125
         Text            =   "Text5"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8640
         TabIndex        =   123
         Text            =   "Text7"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   5
         Left            =   -66480
         TabIndex        =   121
         Text            =   "Text5"
         Top             =   4560
         Width           =   2895
      End
      Begin VB.CommandButton Command10 
         Caption         =   "配比复制"
         Height          =   495
         Left            =   17160
         TabIndex        =   119
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc17 
         Height          =   330
         Left            =   1920
         Top             =   9960
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
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
         Caption         =   "Adodc17"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc16 
         Height          =   330
         Left            =   2280
         Top             =   10080
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1080
         TabIndex        =   118
         Text            =   "Text6"
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "新单据"
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   1200
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc15 
         Height          =   330
         Left            =   7200
         Top             =   10320
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
            Size            =   10.5
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
         Left            =   3000
         Top             =   10080
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Left            =   3720
         Top             =   10200
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
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
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   4
         Left            =   -66480
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   6600
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   3
         Left            =   -66480
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   2
         Left            =   -66480
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   1
         Left            =   -66480
         TabIndex        =   110
         Text            =   "Text5"
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Index           =   0
         Left            =   -66480
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "修改保存"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0FF&
         Caption         =   "添加保存"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2460
         Width           =   1215
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C0C0FF&
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "刷新"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "下一单号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1260
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0FF&
         Caption         =   "打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "下一织号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   900
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   2940
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   3900
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "纱线计划"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "漂染计划"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "订单查询"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3060
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0C0FF&
         Caption         =   "选择打印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   17160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2460
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Forma20.frx":445E
         Left            =   4800
         List            =   "Forma20.frx":446B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   6420
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0C0FF&
         Caption         =   "配比设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   14520
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3060
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   12240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Forma20.frx":4481
         Top             =   3960
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc Adodc12 
         Height          =   330
         Left            =   5760
         Top             =   10140
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
         Caption         =   "Adodc12"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         Left            =   4800
         Top             =   10140
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
         Caption         =   "Adodc11"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Bindings        =   "Forma20.frx":4487
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   12
         Top             =   1740
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "简称"
         Text            =   "DataCombo1"
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4920
         Top             =   10380
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
            Name            =   "宋体"
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
         Bindings        =   "Forma20.frx":449C
         Height          =   4335
         Left            =   240
         TabIndex        =   20
         Top             =   5280
         Width           =   18135
         _cx             =   31988
         _cy             =   7646
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
         GridLines       =   2
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
         FormatString    =   $"Forma20.frx":44B1
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
         OwnerDraw       =   2
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
         Left            =   5400
         Top             =   10260
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
            Name            =   "宋体"
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
         Left            =   5280
         Top             =   10260
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
            Name            =   "宋体"
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
         Left            =   5880
         Top             =   10140
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
            Name            =   "宋体"
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
         Left            =   5880
         Top             =   10140
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
            Name            =   "宋体"
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
         Left            =   5520
         Top             =   10380
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
            Name            =   "宋体"
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
         Left            =   5400
         Top             =   10380
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
            Name            =   "宋体"
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
         Left            =   6000
         Top             =   10260
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
            Name            =   "宋体"
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
         Left            =   6000
         Top             =   10260
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
            Name            =   "宋体"
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
         Left            =   5160
         Top             =   10140
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
            Name            =   "宋体"
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
         Index           =   1
         Left            =   4920
         TabIndex        =   21
         Top             =   1740
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   2
         Left            =   9360
         TabIndex        =   22
         Top             =   5820
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   3
         Left            =   4920
         TabIndex        =   23
         Top             =   2340
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":4586
         Height          =   360
         Index           =   4
         Left            =   4920
         TabIndex        =   24
         Top             =   2940
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "品名"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   5
         Left            =   8640
         TabIndex        =   25
         Top             =   2340
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   6
         Left            =   8640
         TabIndex        =   26
         Top             =   2940
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   7
         Left            =   4920
         TabIndex        =   27
         Top             =   4500
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   8
         Left            =   8640
         TabIndex        =   28
         Top             =   4500
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":459B
         Height          =   360
         Index           =   9
         Left            =   8640
         TabIndex        =   29
         Top             =   3900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "筒颈"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":45B1
         Height          =   360
         Index           =   10
         Left            =   4920
         TabIndex        =   30
         Top             =   3900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "材料名称"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   11
         Left            =   4920
         TabIndex        =   31
         Top             =   7740
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   12
         Left            =   4920
         TabIndex        =   32
         Top             =   8340
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   13
         Left            =   1320
         TabIndex        =   33
         Top             =   7860
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   14
         Left            =   1320
         TabIndex        =   34
         Top             =   8460
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   15
         Left            =   1320
         TabIndex        =   35
         Top             =   9060
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   16
         Left            =   12240
         TabIndex        =   36
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":45C6
         Height          =   360
         Index           =   17
         Left            =   8640
         TabIndex        =   37
         Top             =   3420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "开幅线"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   18
         Left            =   14640
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   19
         Left            =   15960
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":45DC
         Height          =   360
         Index           =   20
         Left            =   12480
         TabIndex        =   40
         Top             =   4740
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "编号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   21
         Left            =   4440
         TabIndex        =   41
         Top             =   7860
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   22
         Left            =   4920
         TabIndex        =   42
         Top             =   7260
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   23
         Left            =   4440
         TabIndex        =   43
         Top             =   9060
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   24
         Left            =   7560
         TabIndex        =   44
         Top             =   8460
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   25
         Left            =   7560
         TabIndex        =   45
         Top             =   7860
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   26
         Left            =   7560
         TabIndex        =   46
         Top             =   8460
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   27
         Left            =   7560
         TabIndex        =   47
         Top             =   9060
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   28
         Left            =   8280
         TabIndex        =   48
         Top             =   8220
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   12240
         TabIndex        =   49
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421440
         Format          =   327876609
         CurrentDate     =   39961
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   12240
         TabIndex        =   50
         Top             =   2880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   8421440
         Format          =   327876609
         CurrentDate     =   39961
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":45F2
         Height          =   360
         Index           =   29
         Left            =   2880
         TabIndex        =   51
         Top             =   6060
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "xm"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   31
         Left            =   11400
         TabIndex        =   52
         Top             =   7500
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":4607
         Height          =   360
         Index           =   32
         Left            =   1080
         TabIndex        =   53
         Top             =   3420
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "业务号"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   33
         Left            =   1080
         TabIndex        =   54
         Top             =   2940
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   34
         Left            =   12120
         TabIndex        =   55
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   35
         Left            =   4920
         TabIndex        =   56
         Top             =   3420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":461D
         Height          =   360
         Index           =   36
         Left            =   1080
         TabIndex        =   57
         Top             =   2340
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "跟单"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   37
         Left            =   4800
         TabIndex        =   58
         Top             =   5940
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   360
         Index           =   38
         Left            =   8640
         TabIndex        =   59
         Top             =   7260
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":4633
         Height          =   360
         Index           =   39
         Left            =   12240
         TabIndex        =   60
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "mc"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Forma20.frx":4648
         Height          =   360
         Index           =   30
         Left            =   -73920
         TabIndex        =   114
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "工艺编号"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "标准织号"
         Height          =   375
         Index           =   19
         Left            =   6240
         TabIndex        =   126
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "委外单价"
         Height          =   375
         Index           =   18
         Left            =   -67440
         TabIndex        =   124
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "匹重"
         Height          =   375
         Index           =   29
         Left            =   7680
         TabIndex        =   122
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "夜班织价"
         Height          =   375
         Index           =   17
         Left            =   -67440
         TabIndex        =   120
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单据号"
         Height          =   375
         Index           =   16
         Left            =   240
         TabIndex        =   117
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工艺编号"
         Height          =   375
         Index           =   7
         Left            =   -74880
         TabIndex        =   115
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   9135
         Left            =   -74880
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   18735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "扣罚"
         Height          =   375
         Index           =   15
         Left            =   -67440
         TabIndex        =   108
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "转数"
         Height          =   375
         Index           =   14
         Left            =   -67440
         TabIndex        =   107
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "质检单价"
         Height          =   375
         Index           =   28
         Left            =   -67440
         TabIndex        =   106
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "白班织价"
         Height          =   375
         Index           =   13
         Left            =   -67440
         TabIndex        =   105
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "发货单价"
         Height          =   375
         Index           =   27
         Left            =   -67440
         TabIndex        =   104
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "序号"
         Height          =   375
         Left            =   7560
         TabIndex        =   103
         Top             =   8220
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单号"
         Height          =   375
         Index           =   6
         Left            =   3960
         TabIndex        =   102
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "款号"
         Height          =   375
         Index           =   4
         Left            =   8520
         TabIndex        =   101
         Top             =   5820
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "织号"
         Height          =   375
         Index           =   3
         Left            =   3960
         TabIndex        =   100
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "欠织"
         Height          =   375
         Index           =   2
         Left            =   6960
         TabIndex        =   99
         Top             =   9060
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "客户"
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   98
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "累计"
         Height          =   375
         Index           =   6
         Left            =   6960
         TabIndex        =   97
         Top             =   8460
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "疵率"
         Height          =   375
         Index           =   8
         Left            =   6960
         TabIndex        =   96
         Top             =   8460
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "疵布"
         Height          =   375
         Index           =   5
         Left            =   6960
         TabIndex        =   95
         Top             =   7860
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "排产"
         Height          =   375
         Index           =   4
         Left            =   3720
         TabIndex        =   94
         Top             =   9060
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "车间"
         Height          =   375
         Index           =   3
         Left            =   11520
         TabIndex        =   93
         Top             =   4740
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "品名"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   92
         Top             =   2940
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "转数"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   91
         Top             =   7260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "幅宽"
         Height          =   375
         Index           =   1
         Left            =   7680
         TabIndex        =   90
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "克重"
         Height          =   375
         Index           =   0
         Left            =   7680
         TabIndex        =   89
         Top             =   2940
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "计划"
         Height          =   375
         Index           =   10
         Left            =   3960
         TabIndex        =   88
         Top             =   4500
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "线长"
         Height          =   375
         Index           =   11
         Left            =   7680
         TabIndex        =   87
         Top             =   4500
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "筒颈"
         Height          =   375
         Index           =   5
         Left            =   7680
         TabIndex        =   86
         Top             =   3900
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "纱支"
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   85
         Top             =   3900
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "针寸"
         Height          =   375
         Index           =   13
         Left            =   10080
         TabIndex        =   84
         Top             =   7860
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "线长"
         Height          =   375
         Index           =   14
         Left            =   600
         TabIndex        =   83
         Top             =   7260
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "配比"
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   82
         Top             =   7860
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "织耗"
         Height          =   375
         Index           =   15
         Left            =   600
         TabIndex        =   81
         Top             =   8460
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "纱量"
         Height          =   375
         Index           =   16
         Left            =   600
         TabIndex        =   80
         Top             =   9060
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "开幅线"
         Height          =   375
         Index           =   8
         Left            =   7680
         TabIndex        =   79
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "下单日期"
         Height          =   375
         Index           =   18
         Left            =   11280
         TabIndex        =   78
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "交期"
         Height          =   375
         Index           =   19
         Left            =   11280
         TabIndex        =   77
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "工艺"
         Height          =   375
         Index           =   9
         Left            =   3720
         TabIndex        =   76
         Top             =   7860
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "业务"
         Height          =   375
         Index           =   10
         Left            =   2040
         TabIndex        =   75
         Top             =   6060
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "提纱"
         Height          =   375
         Left            =   10440
         TabIndex        =   74
         Top             =   7500
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "备注"
         Height          =   375
         Left            =   11280
         TabIndex        =   73
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "类别"
         Height          =   375
         Index           =   11
         Left            =   3960
         TabIndex        =   72
         Top             =   6420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "布类"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   71
         Top             =   3420
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "小备注"
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   70
         Top             =   2940
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "外销号"
         Height          =   375
         Index           =   20
         Left            =   11280
         TabIndex        =   69
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "色别"
         Height          =   375
         Index           =   21
         Left            =   3960
         TabIndex        =   68
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "跟单"
         Height          =   375
         Index           =   12
         Left            =   240
         TabIndex        =   67
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "发货地"
         Height          =   495
         Left            =   11280
         TabIndex        =   66
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "米数"
         Height          =   375
         Index           =   22
         Left            =   3960
         TabIndex        =   65
         Top             =   5940
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单价"
         Height          =   375
         Index           =   23
         Left            =   7680
         TabIndex        =   64
         Top             =   7260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "单位"
         Height          =   375
         Index           =   24
         Left            =   11280
         TabIndex        =   63
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "织价"
         Height          =   375
         Index           =   25
         Left            =   3960
         TabIndex        =   62
         Top             =   7740
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "扣罚"
         Height          =   375
         Index           =   26
         Left            =   3960
         TabIndex        =   61
         Top             =   8340
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "Forma20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strname As String
Dim Stm As New ADODB.Stream
Dim StrPicTemp As String
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private strn As String
Private Sub Command1_Click()
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "01"
Else
Adodc5.RecordSource = "select MAX(right(织号,2)) as h  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc5.Refresh
If Len(Trim(Val(Adodc5.Recordset.Fields(0)) + 1)) = 1 Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "0" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
If Len(Trim(Val(Adodc5.Recordset.Fields(0)) + 1)) = 2 Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
End If

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select  *  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号) from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If

End Sub


Private Sub Command10_Click()
Adodc15.RecordSource = "select 织号 from zbkpd where 单据='" & Text6 & "' and 织号 in(select distinct 织号 sxpb)"
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
MsgBox ("设置配比")
Exit Sub
End If

Adodc15.RecordSource = "select 织号 from sxpb where  织号='" & DataCombo1(3) & "'"
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
sql1 = "insert into  sxpb(织号,排产,纱支,织耗,配比,纱量,批次,备注,序号,颜色,产地) select '" & DataCombo1(3) & "','" & DataCombo1(7) & "',纱支,织耗,配比,纱量,批次,备注,序号,颜色,产地 from sxpb where 织号='" & Text8 & "'"
sql2 = "update  sxpb set 纱量=排产/(100-cast(isnull(织耗,0) as real))*配比 where 织号='" & DataCombo1(3) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
If DataCombo1(4).Text = "" Then
MsgBox ("请输入品名")
Exit Sub
End If



If DataCombo1(7).Text = "" Then
MsgBox ("请输入计划量")
Exit Sub
End If


If DataCombo1(39).Text = "" Then
MsgBox ("请输入单位")
Exit Sub
End If


Adodc15.RecordSource = "select 织号 from zbkpd where 单据='" & Text6 & "'"
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
Adodc15.RecordSource = "select 纱支 from sxpb where 织号='" & DataCombo1(3) & "' order by 序号"
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
MsgBox ("请先设置纱线配比！")
Exit Sub
Else
Adodc15.Recordset.MoveFirst
xxpb = ""
Do While Not Adodc15.Recordset.EOF
xxpb = xxpb + Adodc15.Recordset.Fields(0)
Adodc15.Recordset.MoveNext
Loop
End If

Else

Adodc15.RecordSource = "select 纱支 from sxpb where 织号='" & DataCombo1(3) & "' order by 序号"
Adodc15.Refresh
If Adodc15.Recordset.EOF Then
MsgBox ("请先设置纱线配比！")
Exit Sub
Else
Adodc15.Recordset.MoveFirst
xxpb = ""
Do While Not Adodc15.Recordset.EOF
xxpb = xxpb + Adodc15.Recordset.Fields(0)
Adodc15.Recordset.MoveNext
Loop
End If
End If

Adodc1.Recordset.AddNew
For i = 0 To 9
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
DataCombo1(10).Text = xxpb
Adodc1.Recordset.Fields(10) = DataCombo1(10).Text
Adodc1.Recordset.Fields(11) = DataCombo1(16).Text
Adodc1.Recordset.Fields(12) = DataCombo1(17).Text
Adodc1.Recordset.Fields(13) = DTPicker1.value
Adodc1.Recordset.Fields(14) = DTPicker2.value
Adodc1.Recordset.Fields(15) = DataCombo1(28).Text
Adodc1.Recordset.Fields(16) = DataCombo1(29).Text
Adodc1.Recordset.Fields(17) = DataCombo1(35).Text
Adodc1.Recordset.Fields(18) = DataCombo1(32).Text
Adodc1.Recordset.Fields(19) = DataCombo1(33).Text
Adodc1.Recordset.Fields(20) = DataCombo1(34).Text
Adodc1.Recordset.Fields(21) = DataCombo1(36).Text
Adodc1.Recordset.Fields(22) = DataCombo1(37).Text
Adodc1.Recordset.Fields(23) = DataCombo1(20).Text
For i = 0 To 6
Adodc1.Recordset.Fields(24 + i) = Val(Text5(i))
Next
Adodc1.Recordset.Fields(31) = DataCombo1(30).Text
Adodc1.Recordset.Fields(32) = Text6
Adodc1.Recordset.Fields(33) = Text7
Adodc1.Recordset.Update
Adodc1.Refresh

'Call Command1_Click
Adodc17.Refresh
DataCombo1(7).Text = ""
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号)  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If
Call Command8_Click
End Sub

Private Sub Command12_Click()
Call ddlcd(Adodc8, Adodc12, Text6)
'Call jhd(Adodc8, DataCombo1(3).Text)
End Sub
Private Sub Command15_Click()
'If DataCombo1(1).Text = "" Then
'Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc1.RecordSource = "SELECT 客户,单号,款号,织号,品名,幅宽,克重,计划,纱长,筒颈,纱别,备注,开幅线,日期,交期,序号,业务,颜色,业务号 as 布类,合同号 as 小备注,外销号,跟单,单位,车台,单价,织价,质价,扣罚,转数,夜织,委价,工艺,单据,匹重 FROM zbkpd where 日期=cast('" & DTPicker1.Value & "' as datetime) and left(单号,1)='" & yhdm & "' ORDER BY 单号,织号"
'Adodc1.Refresh
'Else
'Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc1.RecordSource = "SELECT 客户,单号,款号,织号,品名,幅宽,克重,计划,纱长,筒颈,纱别,备注,开幅线,日期,交期,序号,业务,颜色,业务号 as 布类,合同号 as 小备注,外销号,跟单,单位,车台,单价,织价,质价,扣罚,转数,夜织,委价,工艺,单据,匹重 FROM zbkpd where 单号='" & DataCombo1(1).Text & "' ORDER BY 单号,织号"
'Adodc1.Refresh
'End If
'Forma21.Text6 = Text6
'Forma21.Show
End Sub

Private Sub Command16_Click()
If MsgBox("确定配比设置吗？", vbYesNo) = vbNo Then
Forma105.Text1(0).Text = DataCombo1(3).Text
Forma105.Text1(1).Text = DataCombo1(7).Text
Forma105.Text1(2).Text = DataCombo1(10).Text
Forma105.Text2.Text = DataCombo1(1).Text
Forma105.Show
Else
Call Command10_Click
Forma105.Text1(0).Text = DataCombo1(3).Text
Forma105.Text1(1).Text = DataCombo1(7).Text
Forma105.Text1(2).Text = DataCombo1(10).Text
Forma105.Text2.Text = DataCombo1(1).Text
Forma105.Show
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("确定修改吗？", vbYesNo) = vbNo Then Exit Sub

If DataCombo1(4).Text = "" Then
MsgBox ("请输入品名")
Exit Sub
End If

If DataCombo1(7).Text = "" Then
MsgBox ("请输入计划量")
Exit Sub
End If

For i = 0 To 10
Adodc1.Recordset.Fields(i) = DataCombo1(i).Text
Next
Adodc1.Recordset.Fields(11) = DataCombo1(16).Text
Adodc1.Recordset.Fields(12) = DataCombo1(17).Text
Adodc1.Recordset.Fields(13) = DTPicker1.value
Adodc1.Recordset.Fields(14) = DTPicker2.value
Adodc1.Recordset.Fields(15) = DataCombo1(28).Text
Adodc1.Recordset.Fields(16) = DataCombo1(29).Text
Adodc1.Recordset.Fields(17) = DataCombo1(35).Text
Adodc1.Recordset.Fields(18) = DataCombo1(32).Text
Adodc1.Recordset.Fields(19) = DataCombo1(33).Text
Adodc1.Recordset.Fields(20) = DataCombo1(34).Text
Adodc1.Recordset.Fields(21) = DataCombo1(36).Text
Adodc1.Recordset.Fields(22) = DataCombo1(37).Text
Adodc1.Recordset.Fields(23) = DataCombo1(20).Text
For i = 0 To 6
Adodc1.Recordset.Fields(24 + i) = Val(Text5(i))
Next
Adodc1.Recordset.Fields(31) = DataCombo1(30).Text
Adodc1.Recordset.Fields(33) = Text7
Adodc1.Recordset.Update
Adodc1.Refresh

Adodc1.Recordset.Update
Adodc1.Refresh

DataCombo1(7).Text = ""
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号)  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub


Private Sub Command3_Click()
On Error Resume Next
Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.RecordSource = "SELECT * FROM v_zbkpd_djh where 单据编号='" & yhdm & "'"
Adodc16.Refresh

Text6 = Trim(yhdm) + "0000001"
If Adodc16.Recordset.EOF Then
Text6 = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc16.Recordset.Fields(1)) + 1
Text6 = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 客户,单号,款号,织号,品名,幅宽,克重,计划,纱长,筒颈,纱别,备注,开幅线,日期,交期,序号,业务,颜色,业务号 as 布类,合同号 as 小备注,外销号,跟单,单位,车台,单价,织价,质价,扣罚,转数,夜织,委价,工艺,单据,匹重 FROM zbkpd where 单据='" & Text6 & "' ORDER BY 单号,织号"
Adodc1.Refresh
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select  *  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号) from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If
DataCombo1(3) = ""
DataCombo1(1) = ""
DataCombo1(7) = ""
DataCombo1(9) = ""
DataCombo1(5) = ""
DataCombo1(6) = ""
DataCombo1(33) = ""
DataCombo1(34) = ""
Call Command8_Click
Text8 = DataCombo1(3)
End Sub

Private Sub Command33_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
Adodc1.Refresh

sql1 = "delete from sxpb where 织号='" & Adodc1.Recordset.Fields(3) & "'"
sql2 = "delete from zbkpd_jtjh where 织号='" & Adodc1.Recordset.Fields(3) & "'"
sql3 = "delete from clbb where 织号='" & Adodc1.Recordset.Fields(3) & "'"
sql4 = "delete from zjbb where 织号='" & Adodc1.Recordset.Fields(3) & "'"
sql5 = "delete from mprk where 织号='" & Adodc1.Recordset.Fields(3) & "'"
sql6 = "delete from mpck where 织号='" & Adodc1.Recordset.Fields(3) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
RD.Open sql3, conn, adOpenStatic, adLockOptimistic
RD.Open sql4, conn, adOpenStatic, adLockOptimistic
RD.Open sql5, conn, adOpenStatic, adLockOptimistic
RD.Open sql6, conn, adOpenStatic, adLockOptimistic

DataCombo1(7).Text = ""
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号)  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
'Formj4.DataCombo1(0) = DataCombo1(0).Text
'Formj4.DataCombo1(1) = DataCombo1(1).Text
'Formj4.DataCombo1(2) = DataCombo1(2).Text
'Formj4.DataCombo1(10) = DataCombo1(29).Text
'Formj4.Show
End Sub

Private Sub Command6_Click()
'Formj1.DataCombo1(1) = DataCombo1(1).Text
'Formj1.DataCombo1(2) = DataCombo1(2).Text
'Formj1.DataCombo1(14) = DataCombo1(29).Text
'Formj1.DataCombo1(2) = DataCombo1(3).Text
'Formj1.DataCombo1(3) = DataCombo1(4).Text
'Formj1.DataCombo1(4) = DataCombo1(5).Text
'Formj1.DataCombo1(5) = DataCombo1(6).Text
'Formj1.DataCombo1(6) = DataCombo1(7).Text
'Formj1.DataCombo1(7) = DataCombo1(35).Text
'Formj1.Show
End Sub

Private Sub Command7_Click()
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 客户,单号,款号,织号,品名,幅宽,克重,计划,纱长,筒颈,纱别,备注,开幅线,日期,交期,序号,业务,颜色,业务号 as 布类,合同号 as 小备注,外销号,跟单,单位,车台,单价,织价,质价,扣罚,转数,夜织,委价,工艺,单据,匹重 FROM zbkpd where 单据='" & Text6 & "' ORDER BY 单号,织号"
Adodc1.Refresh
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select  *  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号) from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If
Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command8_Click()
'If Combo1.Text = "加工" Then
'Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.Value & "' as datetime) and left(单号,1)='" & yhdm & "'"
'Adodc4.Refresh
'If Adodc4.Recordset.EOF Then
'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + "01J"
'Else
'Adodc4.RecordSource = "select MAX(LEFT(right(单号,3),2)) as h  from zbkpd where 日期=cast('" & DTPicker1.Value & "' as datetime) and left(单号,1)='" & yhdm & "'"
'Adodc4.Refresh
'If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 1 Then
'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + "0" + Trim(Adodc4.Recordset.Fields(0) + 1) + "J"
'End If
'If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 2 Then
'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + Trim(Adodc4.Recordset.Fields(0) + 1) + "J"
'End If
'End If
'End If


'If Combo1.Text = "外协" Then
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(单号,1)='" & yhdm & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "01"
Else
Adodc4.RecordSource = "select MAX(right(单号,2)) as h  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(单号,1)='" & yhdm & "'"
Adodc4.Refresh
If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 1 Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "0" + Trim(Adodc4.Recordset.Fields(0) + 1)
End If
If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 2 Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + Trim(Adodc4.Recordset.Fields(0) + 1)
End If
End If
'End If

'If Combo1.Text = "销售" Then
'Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
'Adodc4.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.Value & "' as datetime) and left(单号,1)='" & yhdm & "'"
'Adodc4.Refresh
'If Adodc4.Recordset.EOF Then

'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + "01X"
'Else
'Adodc4.RecordSource = "select MAX(LEFT(right(单号,3),2)) as h  from zbkpd where 日期=cast('" & DTPicker1.Value & "' as datetime) and left(单号,1)='" & yhdm & "'"
'Adodc4.Refresh
'If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 1 Then
'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + "0" + Trim(Adodc4.Recordset.Fields(0) + 1) + "X"
'End If
'If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 2 Then
'DataCombo1(1).Text = yhdm + Format(DTPicker1.Value, "YYMMDD") + Trim(Adodc4.Recordset.Fields(0) + 1) + "X"
'End If
'End If
'End If
DataCombo1(3).Text = DataCombo1(1)
DataCombo1(7) = ""
DataCombo1(9) = ""
DataCombo1(5) = ""
DataCombo1(6) = ""
DataCombo1(33) = ""
DataCombo1(34) = ""
End Sub



Private Sub Command9_Click()
'Formj31.DataCombo6 = DataCombo1(1).Text
'Formj31.Show
End Sub

Private Sub DataCombo1_Click(Index As Integer, Area As Integer)
On Error Resume Next
Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc13.RecordSource = "SELECT * FROM gytp where 工艺编号='" & DataCombo1(30).Text & "'"
Adodc13.Refresh

       
    Image1.Picture = Nothing

If Adodc13.Recordset.Fields(3).Type = 205 Then
     StrPicTemp = "c:\temp.tmp"     '临时文件,用来保存读出的图片
     With Stm
        .Type = adTypeBinary
        .Open
        .Write Adodc13.Recordset.Fields(3)        '写入数据库中的数据至Stream中
        .SaveToFile StrPicTemp, adSaveCreateOverWrite   '将Stream中数据写入临时文件中
        .Close
    End With
    
    Image1.Picture = LoadPicture(StrPicTemp)

End If
End Sub

Private Sub dataCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Form_Load()
On Error Resume Next
For i = 0 To 39
DataCombo1(i).Text = ""
Next
Combo1.Text = "加工"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Text8.Text = ""
For i = 0 To 6
Text5(i) = 0
Next
Text5(0) = 5
Text5(1) = 0.5
Text5(5) = 0.6
Text5(6) = 3
Text5(2) = 0.03
Text5(3) = 0.5
Text5(4) = 100
Text6 = ""
DTPicker1.value = Date
DTPicker2.value = Date
DataCombo1(39).Text = "公斤"

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(单号,1)='" & yhdm & "'"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "01"
Else
Adodc4.RecordSource = "select MAX(right(单号,2)) as h  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(单号,1)='" & yhdm & "'"
Adodc4.Refresh
If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 1 Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "0" + Trim(Adodc4.Recordset.Fields(0) + 1)
End If
If Len(Trim(Val(Adodc4.Recordset.Fields(0)) + 1)) = 2 Then
DataCombo1(1).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + Trim(Adodc4.Recordset.Fields(0) + 1)
End If
End If

Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc5.RecordSource = "select *  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc5.Refresh
If Adodc5.Recordset.EOF Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "01"
Else
Adodc5.RecordSource = "select MAX(right(织号,2)) as h  from zbkpd where 日期=cast('" & DTPicker1.value & "' as datetime) and left(织号,1)='" & yhdm & "'"
Adodc5.Refresh
If Len(Trim(Val(Adodc5.Recordset.Fields(0)) + 1)) = 1 Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + "0" + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
If Len(Trim(Val(Adodc5.Recordset.Fields(0)) + 1)) = 2 Then
DataCombo1(3).Text = yhdm + Format(DTPicker1.value, "YYMMDD") + Trim(Adodc5.Recordset.Fields(0) + 1)
End If
End If


Adodc7.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc7.RecordSource = "select xm from ywf group by xm"
Adodc7.Refresh

Adodc9.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc9.RecordSource = "select MC from CLDW group by MC"
Adodc9.Refresh

Adodc10.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc10.RecordSource = "select distinct 开幅线 from zbkpd"
Adodc10.Refresh

Adodc11.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc11.RecordSource = "select distinct 筒颈 from zbkpd"
Adodc11.Refresh

Adodc12.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc12.RecordSource = "SELECT 编号 FROM CT  GROUP BY 编号"
Adodc12.Refresh

Adodc13.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc14.RecordSource = "SELECT 工艺编号 FROM gytp  GROUP BY 工艺编号"
Adodc14.Refresh

Adodc17.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc17.RecordSource = "SELECT distinct 业务号 FROM zbkpd"
Adodc17.Refresh

Adodc18.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc18.RecordSource = "SELECT distinct 跟单 FROM zbkpd"
Adodc18.Refresh


Adodc16.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc16.RecordSource = "SELECT * FROM v_zbkpd_djh where 单据编号='" & yhdm & "'"
Adodc16.Refresh

Text6 = Trim(yhdm) + "0000001"
If Adodc16.Recordset.EOF Then
Text6 = Trim(yhdm) + "0000001"
Else
uu = Val(Adodc16.Recordset.Fields(1)) + 1
Text6 = Trim(yhdm) + Left("0000000", 7 - Len(uu)) + Trim(uu)
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT 客户,单号,款号,织号,品名,幅宽,克重,计划,纱长,筒颈,纱别,备注,开幅线,日期,交期,序号,业务,颜色,业务号 as 布类,合同号 as 小备注,外销号,跟单,单位,车台,单价,织价,质价,扣罚,转数,夜织,委价,工艺,单据,匹重 FROM zbkpd where 单据='" & Text6 & "' ORDER BY 单号,织号"
Adodc1.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select  *  from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
If Adodc6.Recordset.EOF Then
DataCombo1(28).Text = 1
Else
Adodc6.RecordSource = "select MAX(序号) from zbkpd where 单据='" & Text6 & "'"
Adodc6.Refresh
DataCombo1(28).Text = Adodc6.Recordset.Fields(0) + 1
End If

Adodc15.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
VSFlexGrid1.ColWidth(i) = 1500
Next
VSFlexGrid1.ColWidth(24) = 0

Command11.Enabled = True
Command2.Enabled = False
Command4.Enabled = False

End Sub

Private Sub Label3_Click(Index As Integer)
Select Case Index
       Case 6
DataCombo1(1).Enabled = True
       Case 3
DataCombo1(3).Enabled = True
End Select
End Sub

Private Sub Label3_DblClick(Index As Integer)
Select Case Index
       Case 6
DataCombo1(1).Enabled = False
       Case 3
DataCombo1(3).Enabled = False
End Select
End Sub

Private Sub Label5_Click()
beizhu = 11
Forma112.Text1(3) = DataCombo1(16).Text
Forma112.Show
End Sub

Private Sub Text1_Change()
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select 简称 from v_khZL where 简码 like '%'+'" & Text1.Text & "'+'%' group by 简称"
Adodc2.Refresh
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_Change()
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select 品名 from v_zbkpd_pm where 简码 like '%'+'" & Text2.Text & "'+'%' group by 品名"
Adodc3.Refresh
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text3_Change()
'On Error Resume Next
Adodc8.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc8.RecordSource = "select 材料名称 from clmc where 材料名称 like '%'+'" & Text3.Text & "'+'%' group by 材料名称"
Adodc8.Refresh
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode

End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1


For i = 0 To 10
 DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
DataCombo1(16).Text = Adodc1.Recordset.Fields(11)
DataCombo1(17).Text = Adodc1.Recordset.Fields(12)
DTPicker1.value = Adodc1.Recordset.Fields(13)
DTPicker2.value = Adodc1.Recordset.Fields(14)
DataCombo1(28).Text = Adodc1.Recordset.Fields(15)
DataCombo1(29).Text = Adodc1.Recordset.Fields(16)
 DataCombo1(35).Text = Adodc1.Recordset.Fields(17)
DataCombo1(32).Text = Adodc1.Recordset.Fields(18)
DataCombo1(33).Text = Adodc1.Recordset.Fields(19)
 DataCombo1(34).Text = Adodc1.Recordset.Fields(20)

 DataCombo1(36).Text = Adodc1.Recordset.Fields(21)
DataCombo1(37).Text = Adodc1.Recordset.Fields(22)
DataCombo1(38).Text = Adodc1.Recordset.Fields(23)

For i = 0 To 6
Text5(i) = Adodc1.Recordset.Fields(24 + i)
Next
DataCombo1(30).Text = Adodc1.Recordset.Fields(31)
Text7 = Adodc1.Recordset.Fields(31)

If Adodc1.Recordset.Fields(17) <> "" And Adodc1.Recordset.Fields(17) <> Null Then
Command11.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Exit Sub
End If

If Adodc1.Recordset.Fields(18) > 0 Then
Command11.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Exit Sub
End If

For i = 0 To 10
 DataCombo1(i).Text = Adodc1.Recordset.Fields(i)
Next
DataCombo1(16).Text = Adodc1.Recordset.Fields(11)
DataCombo1(17).Text = Adodc1.Recordset.Fields(12)
DTPicker1.value = Adodc1.Recordset.Fields(13)
DTPicker2.value = Adodc1.Recordset.Fields(14)
DataCombo1(28).Text = Adodc1.Recordset.Fields(15)
DataCombo1(29).Text = Adodc1.Recordset.Fields(16)
DataCombo1(35).Text = Adodc1.Recordset.Fields(17)
DataCombo1(32).Text = Adodc1.Recordset.Fields(18)
DataCombo1(33).Text = Adodc1.Recordset.Fields(19)
DataCombo1(34).Text = Adodc1.Recordset.Fields(20)
 
DataCombo1(36).Text = Adodc1.Recordset.Fields(21)
DataCombo1(37).Text = Adodc1.Recordset.Fields(22)
DataCombo1(38).Text = Adodc1.Recordset.Fields(23)

For i = 0 To 6
Text5(i) = Adodc1.Recordset.Fields(24 + i)
Next
DataCombo1(30).Text = Adodc1.Recordset.Fields(31)
Text7 = Adodc1.Recordset.Fields(31)

Command11.Enabled = False
Command2.Enabled = True
Command4.Enabled = True
End Sub
