VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Forma902 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ⱦ��Ԥ����Ϣ"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12720
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   9000
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9600
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
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
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "��ʾ��Ϣ"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   8175
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���ڿ������"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�����������"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���������֮��"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ȫ��"
         Height          =   255
         Left            =   6000
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Forma902.frx":0000
      Height          =   450
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   794
      _Version        =   393216
      ListField       =   "����"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1560
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
         Name            =   "����"
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
      Bindings        =   "Forma902.frx":0015
      Height          =   5535
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   11655
      _cx             =   20558
      _cy             =   9763
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   7320
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Ⱦ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Forma902"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sj As Integer
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Call MXOutadodcToExcel(VSFlexGrid1, "�����Ϣ")
End Sub

Private Sub Command6_Click()
On Error Resume Next
If Option1.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_rlts  where ���>=������� order by ����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_rlts  where  ���>=������� and Ⱦ�� like '%'+'" & DataCombo1.Text & "'+'%' order by ����"
Adodc1.Refresh
End If
End If
If Option2.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_rlts  where ���<������� and  ���>������� order by ����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_rlts  where  ���<������� and  ���>�������  and Ⱦ�� like '%'+'" & DataCombo1.Text & "'+'%' order by ����"
Adodc1.Refresh
End If
End If
If Option3.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_rlts  order by ����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_rlts  where Ⱦ�� like '%'+'" & DataCombo1.Text & "'+'%' order by ����"
Adodc1.Refresh
End If
End If
If Option4.value = True Then
If DataCombo1.Text = "" Then
Adodc1.RecordSource = "SELECT * from yj_rlts  where ���<������� order by ����"
Adodc1.Refresh
Else
Adodc1.RecordSource = "SELECT * from yj_rlts  where  ���<������� and Ⱦ�� like '%'+'" & DataCombo1.Text & "'+'%' order by ����"
Adodc1.Refresh
End If
End If

VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(3) = 2000
VSFlexGrid1.ColWidth(4) = 2000
VSFlexGrid1.ColWidth(5) = 2000

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Form_Load()

Text2.Text = ""
DataCombo1.Text = ""
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "SELECT * from yj_rlts  where ���>=������� order by ����"
Adodc1.Refresh
Option4.value = True
VSFlexGrid1.ColWidth(1) = 2000
VSFlexGrid1.ColWidth(2) = 2000
VSFlexGrid1.ColWidth(3) = 2000
VSFlexGrid1.ColWidth(4) = 2000
VSFlexGrid1.ColWidth(5) = 2000

If VSFlexGrid1.Rows > 1 Then
For i = 1 To VSFlexGrid1.Rows - 1
VSFlexGrid1.RowHeight(i) = 600
Next
End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState = 1 Then
  If Not Adodc1.Recordset.EOF Then
    Timer2.Enabled = True
  Else
    Timer2.Enabled = False
  End If
  End If
  If Me.WindowState = 0 Then
    Timer2.Enabled = False
  End If
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then Exit Sub
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.RecordSource = "select distinct ���� from rhzh where ���� like '%'+'" & Text2.Text & "'+'%'  order by ����"
Adodc2.Refresh
End Sub

Private Sub Timer1_Timer()
If sj = 5 Then
sj = 1
Call Command6_Click
Else
sj = sj + 1
End If
End Sub

Private Sub Timer2_Timer()
    Dim i
    i = FlashWindow(Me.hwnd, 1) '��ʱ��˸������
End Sub
