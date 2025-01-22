VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma110 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ã«Å÷Âëµ¥²Ù×÷"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   11445
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   2160
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   480
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "±êÇ©´òÓ¡"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Âëµ¥´òÓ¡"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ÍË³ö"
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   840
      Top             =   8400
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
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1320
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1080
      Top             =   8280
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
         Name            =   "ËÎÌå"
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
      Left            =   2040
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
         Name            =   "ËÎÌå"
         Size            =   9
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
      Caption         =   "É¾³ý"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Â¼Èë"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ÐÞ¸Ä"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ë¢ÐÂ"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   1575
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
      Bindings        =   "Forma110.frx":0000
      Height          =   3855
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   5175
      _cx             =   9128
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
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
      Bindings        =   "Forma110.frx":0015
      Height          =   7215
      Left            =   6000
      TabIndex        =   4
      Top             =   840
      Width           =   4695
      _cx             =   8281
      _cy             =   12726
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   2160
      Top             =   8280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
         Name            =   "ËÎÌå"
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
      Left            =   2280
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
         Name            =   "ËÎÌå"
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
      Left            =   1680
      Top             =   8280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÐòºÅ"
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Èë¿âµ¥¾Ý"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   20
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÖØÁ¿"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Æ¥ºÅ"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¹øºÅ"
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Æ·Ãû"
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ã«Åß·ù¿í"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Forma110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select isnull(count(Æ¥ºÅ),0),isnull(sum(ÖØÁ¿),0) from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' "
Adodc4.Refresh

If Val(Forma103.DataCombo2(5)) > Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
End If

If Val(Forma103.DataCombo2(5)) = Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
MsgBox ("Ã«Å÷Âëµ¥ÒÑ·ûºÏ³ö¿âÆ¥Êý")
Exit Sub
Me.Hide
End If

If Val(Forma103.DataCombo2(5)) < Val(Adodc4.Recordset.Fields(0)) Then
Exit Sub
Me.Hide
End If

Adodc3.Recordset.AddNew
For i = 0 To 6
Adodc3.Recordset.Fields(i) = Text1(i).Text
Next
Adodc3.Recordset.Update
Adodc3.Refresh

Adodc2.RecordSource = "select Æ¥ºÅ from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Text1(3).Text = 1
Else
Text1(3).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text1(4).Text = ""
Text1(4).SetFocus
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()
On Error Resume Next
If MsgBox("È·¶¨É¾³ýÂð£¿", vbYesNo) = vbNo Then Exit Sub
Adodc3.Recordset.Delete
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select isnull(count(Æ¥ºÅ),0),isnull(sum(ÖØÁ¿),0) from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' "
Adodc4.Refresh
If Val(Forma103.DataCombo2(5)) > Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
End If

If Val(Forma103.DataCombo2(5)) = Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
MsgBox ("Ã«Å÷Âëµ¥ÒÑ·ûºÏ³ö¿âÆ¥Êý")
Exit Sub
Me.Hide
End If

If Val(Forma103.DataCombo2(5)) < Val(Adodc4.Recordset.Fields(0)) Then
Exit Sub
Me.Hide
End If

Adodc2.RecordSource = "select Æ¥ºÅ from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Text1(3).Text = 1
Else
Text1(3).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text1(4).Text = ""
Command1.Enabled = True
Command9.Enabled = False
Command3.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command4_Click()
Call lcd3(Adodc5, Adodc6, Text1(0).Text, Text1(6).Text)
End Sub

Private Sub Command6_Click()
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select isnull(count(Æ¥ºÅ),0),isnull(sum(ÖØÁ¿),0) from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' "
Adodc4.Refresh
If Val(Forma103.DataCombo2(5)) > Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
End If

If Val(Forma103.DataCombo2(5)) = Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
MsgBox ("Ã«Å÷Âëµ¥ÒÑ·ûºÏ³ö¿âÆ¥Êý")
Exit Sub
Me.Hide
End If

If Val(Forma103.DataCombo2(5)) < Val(Adodc4.Recordset.Fields(0)) Then
Exit Sub
Me.Hide
End If

Adodc1.RecordSource = "select Æ·Ãû,Ã«Åß·ù¿í from kpd where ¹øºÅ='" & Text1(0).Text & "'"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select Æ¥ºÅ from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Text1(3).Text = 1
Else
Text1(3).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text1(4).Text = ""
Adodc3.RecordSource = "select * from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc3.Refresh
Command1.Enabled = True
Command9.Enabled = False
Command3.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Command9_Click()
On Error Resume Next
If MsgBox("È·¶¨ÐÞ¸ÄÂð£¿", vbYesNo) = vbNo Then Exit Sub
For i = 0 To 6
Adodc3.Recordset.Fields(i) = Text1(i).Text
Next
Adodc3.Recordset.Update
Adodc3.Refresh

Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select isnull(count(Æ¥ºÅ),0),isnull(sum(ÖØÁ¿),0) from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' "
Adodc4.Refresh
If Val(Forma103.DataCombo2(5)) > Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
End If

If Val(Forma103.DataCombo2(5)) = Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
MsgBox ("Ã«Å÷Âëµ¥ÒÑ·ûºÏ³ö¿âÆ¥Êý")
Exit Sub
Me.Hide
End If

If Val(Forma103.DataCombo2(5)) < Val(Adodc4.Recordset.Fields(0)) Then
Exit Sub
Me.Hide
End If

Adodc2.RecordSource = "select Æ¥ºÅ from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Text1(3).Text = 1
Else
Text1(3).Text = Adodc2.Recordset.Fields(0) + 1
End If
Text1(4).Text = ""
Command1.Enabled = True
Command9.Enabled = False
Command3.Enabled = False
Text1(4).SetFocus
End Sub

Private Sub Form_Load()
For i = 0 To 6
Text1(i).Text = ""
Next
Adodc5.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Command1.Enabled = True
Command9.Enabled = False
Command3.Enabled = False
VSFlexGrid1.ColWidth(0) = 300
VSFlexGrid1.ColWidth(1) = 0
VSFlexGrid1.ColWidth(2) = 0
VSFlexGrid1.ColWidth(3) = 0
VSFlexGrid1.ColWidth(6) = 1500
VSFlexGrid3.ColWidth(1) = 2000
VSFlexGrid3.ColWidth(2) = 1000
End Sub

Private Sub Text1_Change(Index As Integer)
On Error Resume Next
Select Case Index
       Case 6
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select isnull(count(Æ¥ºÅ),0),isnull(sum(ÖØÁ¿),0) from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' "
Adodc4.Refresh
If Val(Forma103.DataCombo2(5)) > Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
End If

If Val(Forma103.DataCombo2(5)) = Val(Adodc4.Recordset.Fields(0)) Then
Forma103.DataCombo2(6) = Adodc4.Recordset.Fields(1)
'Exit Sub
Me.Hide
End If

If Val(Forma103.DataCombo2(5)) < Val(Adodc4.Recordset.Fields(0)) Then
MsgBox ("Ã«Å÷Âëµ¥ÒÑ·ûºÏ³ö¿âÆ¥Êý")
'Exit Sub
Me.Hide
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select Æ·Ãû,Ã«Åß·ù¿í from kpd where ¹øºÅ='" & Text1(0).Text & "'"
Adodc1.Refresh
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select Æ¥ºÅ from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
Text1(3).Text = 1
Else
Text1(3).Text = Adodc2.Recordset.Fields(0) + 1
End If
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.RecordSource = "select * from mpbmd where ¹øºÅ='" & Text1(0).Text & "' and ÐòºÅ='" & Text1(6) & "' order by Æ¥ºÅ desc"
Adodc3.Refresh
End Select
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
entertotab KeyCode
End Sub

Private Sub VSFlexGrid1_Click()
On Error Resume Next
If Adodc3.Recordset.EOF Then Exit Sub
Adodc3.Recordset.MoveFirst
rs = VSFlexGrid1.Row
Adodc3.Recordset.Move rs - 1
For i = 0 To 6
Text1(i) = Adodc3.Recordset.Fields(i)
Next
Command1.Enabled = False
Command9.Enabled = True
Command3.Enabled = True
End Sub

Private Sub VSFlexGrid3_dblClick()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
rs = VSFlexGrid3.Row
Adodc1.Recordset.Move rs - 1
For i = 0 To 1
Text1(i + 1).Text = Adodc1.Recordset.Fields(i)
Next
End Sub
