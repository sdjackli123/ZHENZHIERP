VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formw41 
   BackColor       =   &H00C0E0FF&
   Caption         =   "日期设定"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form41"
   ScaleHeight     =   8460
   ScaleWidth      =   7845
   StartUpPosition =   2  '屏幕中心
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   7800
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
      Bindings        =   "Formw41.frx":0000
      Height          =   5175
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   6975
      _cx             =   12303
      _cy             =   9128
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
      AllowUserResizing=   0
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
      AutoSizeMode    =   0
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
      WordWrap        =   0   'False
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5880
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "返回"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "删除"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "修改保存"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "添加保存"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/M/d"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   327876609
      CurrentDate     =   36892
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      CalendarTitleBackColor=   8421440
      CalendarTrailingForeColor=   255
      Format          =   327876609
      CurrentDate     =   36892
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
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   840
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
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Formw41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BA As Database: Dim RD As Recordset
Private Sub Command6_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text1(0).Text = DTPicker1.value
End Sub

Private Sub DTPicker1_CloseUp()
Text1(0).Text = DTPicker1.value
Text1(0).SetFocus
End Sub

Private Sub DTPicker2_Change()
Text1(1).Text = DTPicker2.value
End Sub

Private Sub DTPicker2_CloseUp()
Text1(1).Text = DTPicker2.value
Text1(2).SetFocus
End Sub

Private Sub Form_Load()

On Error Resume Next

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "rqsd"
Adodc1.Refresh
VSFlexGrid1.ColWidth(1) = 1400
VSFlexGrid1.ColWidth(2) = 1400
VSFlexGrid1.ColWidth(3) = 2400

For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next

DTPicker1.value = Date
DTPicker2.value = Date
Text1(0).Text = Date
Text1(1).Text = Date

End Sub






Private Sub Command1_Click()
On Error Resume Next
If Text1(0).Text = "" Or Text1(1).Text = "" Then
MsgBox ("日期输入完整")
Exit Sub
End If
Adodc1.Recordset.AddNew
For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next

End Sub

Private Sub Command2_Click()
On Error Resume Next

For i = 0 To Adodc1.Recordset.Fields.count - 1
Adodc1.Recordset.Fields(i) = Text1(i).Text
Next
Adodc1.Recordset.Update
Adodc1.Refresh
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next

End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Refresh
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = ""
Next

End Sub


Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
For i = 0 To Adodc1.Recordset.Fields.count - 1
Text1(i).Text = Adodc1.Recordset.Fields(i)
Next

End Sub

