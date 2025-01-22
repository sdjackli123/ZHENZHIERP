VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formm3 
   BackColor       =   &H00C0E0FF&
   Caption         =   "权限设置"
   ClientHeight    =   10215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
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
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formm3.frx":0000
      Height          =   5415
      Left            =   600
      TabIndex        =   24
      Top             =   3960
      Width           =   5775
      _cx             =   10186
      _cy             =   9551
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
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
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
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
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
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   1095
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
      Height          =   495
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
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
      Height          =   495
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "录入"
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   7200
      Left            =   6840
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2040
      Width           =   7095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8760
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8880
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   8520
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8280
      Top             =   10800
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
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "模块"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   23
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "全清"
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "全选"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "用户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6840
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "序号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "用户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "代码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "用户信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Formm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Dim conn1 As ADODB.Connection: Dim RD1 As ADODB.Recordset
Dim conn2 As ADODB.Connection: Dim RD2 As ADODB.Recordset
Private Sub Command5_Click()
On Error Resume Next
For i = 0 To List1.ListCount - 1
l1 = Mid(List1.List(i), 1, InStr(List1.List(i), "/") - 1)
L2 = Mid(List1.List(i), InStr(List1.List(i), "/") + 1)
If List1.Selected(i) = True Then
sql1 = "INSERT INTO QXB(用户,功能,代号,权限) VALUES('" & Text5.Text & "','" & l1 & "','" & L2 & "','Y')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "INSERT INTO QXB(用户,功能,代号,权限) VALUES('" & Text5.Text & "','" & l1 & "','" & L2 & "','N')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
List1.Clear
End Sub

Private Sub Command6_Click()
On Error Resume Next
For i = 0 To 200
qxsz(i) = "N"
Next

List1.Clear


If Text5.Text = "" Then
MsgBox ("请输入用户")
Exit Sub
End If
Adodc2.RecordSource = "select * from yhb where 用户='" & Text5.Text & "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
MsgBox ("没有此用户")
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Exit Sub
End If
Adodc3.RecordSource = "select * from qxb where 用户='" & Text5.Text & "' order by 代号"
Adodc3.Refresh

Adodc4.RecordSource = "select * from qxb where 用户='" & Text5.Text & "'"
Adodc4.Refresh


If Adodc3.Recordset.EOF Then
Command5.Enabled = True
Command7.Enabled = False
Command8.Enabled = False
For i = 0 To 200
List1.AddItem Formm1.Label1(i).Caption + "/" + Trim(i)
Next
Else
Command5.Enabled = False
Command7.Enabled = True
Command8.Enabled = True
Adodc3.Recordset.MoveFirst
i = 0
Do While Not Adodc3.Recordset.EOF
List1.AddItem Adodc3.Recordset.Fields(1) + "/" + Adodc3.Recordset.Fields(2)
If Adodc3.Recordset.Fields(3) = "Y" Then
qxsz(i) = "Y"
Else
qxsz(i) = "N"
End If
i = i + 1
Adodc3.Recordset.MoveNext
Loop

For i = 0 To 200
If qxsz(i) = "Y" Then
List1.Selected(i) = True
End If
Next

End If

End Sub

Private Sub Command7_Click()
On Error Resume Next
For i = 0 To List1.ListCount - 1
l1 = Mid(List1.List(i), 1, InStr(List1.List(i), "/") - 1)
L2 = Mid(List1.List(i), InStr(List1.List(i), "/") + 1)
sql1 = "DELETE  FROM QXB WHERE 用户='" & Text5.Text & "' and 代号='" & L2 & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
If List1.Selected(i) = True Then
sql1 = "INSERT INTO QXB(用户,功能,代号,权限) VALUES('" & Text5.Text & "','" & l1 & "','" & L2 & "','Y')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
Else
sql1 = "INSERT INTO QXB(用户,功能,代号,权限) VALUES('" & Text5.Text & "','" & l1 & "','" & L2 & "','N')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
List1.Clear
End Sub

Private Sub Command8_Click()
If MsgBox("确定删除吗？", vbYesNo) = vbNo Then Exit Sub
sql1 = "DELETE  FROM QXB WHERE 用户='" & Text5.Text & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
List1.Clear
End Sub

Private Sub Command9_Click()
If MsgBox("确定转入吗？", vbYesNo) = vbNo Then Exit Sub

sql1 = "delete  from yhb"
sql2 = "delete  from qxb"
RD1.Open sql1, conn1, adOpenStatic, adLockOptimistic
RD1.Open sql2, conn1, adOpenStatic, adLockOptimistic

sql1 = "insert into zzpr.dbo.yhb  select * from zzpr.dbo.yhb"
sql2 = "insert into zzpr.dbo.qxb  select * from zzpr.dbo.qxb"
RD2.Open sql1, conn2, adOpenStatic, adLockOptimistic
RD2.Open sql2, conn2, adOpenStatic, adLockOptimistic

MsgBox ("转入成功！")
End Sub

Private Sub Label5_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Label6_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
rs = VSFlexGrid1.Row
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Text1.Text = Adodc1.Recordset.Fields(0)
Text2.Text = Adodc1.Recordset.Fields(1)
Text5.Text = Adodc1.Recordset.Fields(0)
Text3.Text = Adodc1.Recordset.Fields(2)
Text4.Text = Adodc1.Recordset.Fields(3)
Text6.Text = Adodc1.Recordset.Fields(4)
Text7.Text = Adodc1.Recordset.Fields(5)
Text8.Text = Adodc1.Recordset.Fields(6)
End Sub

Private Sub Command3_Click()
On Error Resume Next

Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Text3.Text
Adodc1.Recordset.Fields(3) = Text4.Text
Adodc1.Recordset.Fields(4) = Text6.Text
Adodc1.Recordset.Fields(5) = Text7.Text
Adodc1.Recordset.Fields(6) = Text8.Text
Adodc1.Recordset.Update
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text6.Text = ""
Text4.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7 = ""
Text4.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next

Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Set conn1 = New ADODB.Connection
conn1.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD1 = New ADODB.Recordset

Set conn2 = New ADODB.Connection
conn2.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD2 = New ADODB.Recordset


Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from YHB"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text4.Text = Adodc1.Recordset.RecordCount + 1
VSFlexGrid1.ColWidth(0) = 200
Command5.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Text1.TabIndex = 0
End Sub
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = Text1.Text
Adodc1.Recordset.Fields(1) = Text2.Text
Adodc1.Recordset.Fields(2) = Text3.Text
Adodc1.Recordset.Fields(3) = Text4.Text
Adodc1.Recordset.Fields(4) = Text6.Text
Adodc1.Recordset.Fields(5) = Text7.Text
Adodc1.Recordset.Fields(6) = Text8.Text
Adodc1.Recordset.Update
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text6.Text = ""
Text4.Text = Adodc1.Recordset.RecordCount + 1
Text1.SetFocus
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    entertotab KeyCode
End Sub

