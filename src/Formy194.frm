VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formy194 
   BackColor       =   &H00C0E0FF&
   Caption         =   "裁剪查询"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12375
   LinkTopic       =   "Form41"
   ScaleHeight     =   9855
   ScaleWidth      =   12375
   StartUpPosition =   3  '窗口缺省
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
      Top             =   9000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "款号刷新"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Formy194.frx":0000
      Height          =   7215
      Left            =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12726
      _Version        =   393216
      Cols            =   10
      BackColorFixed  =   9803263
      BackColorBkg    =   42662
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80936961
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   80936961
      CurrentDate     =   36892
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "款号"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "起始日期："
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
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "结束日期："
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
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Formy194"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call WXCX(MSFlexGrid1, "裁剪查询")
End Sub

Private Sub Command2_Click()
If Text2.Text = "" Then
Data2.RecordSource = "select 款号,颜色,规格,计划,sum(val(裁剪)) as 裁剪量 from cjrb  group by 款号,颜色,规格,计划 order by 款号,颜色,规格,计划"
Data2.Refresh
Else
Data2.RecordSource = "select 款号,颜色,规格,计划,sum(val(裁剪)) as 裁剪量 from cjrb where 款号='" & Text2.Text & "' group by 款号,颜色,规格,计划 order by 款号,颜色,规格,计划"
Data2.Refresh
End If
Call sx
End Sub
Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Form_Load()
Text2.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
Data1.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"

Data2.DatabaseName = "d:\数据库\\htgl\2011\scjd.mdb"
MSFlexGrid1.ColWidth(0) = 300
For i = 1 To 5
MSFlexGrid1.ColWidth(i) = 1200
Next

End Sub

Private Sub Label1_Click(Index As Integer)
Select Case Index
       Case 1
       khbl = 17
Formy202.Show
End Select
End Sub

Private Sub sx()

    Dim i     As Integer
      With MSFlexGrid1
                 .AllowBigSelection = True           '   设置网格样式
                 .FillStyle = flexFillRepeat
                For i = 1 To .Rows - 1
                        .Row = i:       .Col = .FixedCols
                        .ColSel = .Cols() - .FixedCols - 1
                         If Val(MSFlexGrid1.TextMatrix(i, 4)) <= Val(MSFlexGrid1.TextMatrix(i, 5)) Then
                              .CellBackColor = vbBlack           '   黑色
                        Else
                              .CellBackColor = vbGreen      '   兰色
                        End If
                Next i
        End With
End Sub


