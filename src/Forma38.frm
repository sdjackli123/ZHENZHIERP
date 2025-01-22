VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forma38 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form38"
   MaxButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   7905
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1680
      Top             =   6840
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
      Bindings        =   "Forma38.frx":0000
      Height          =   6855
      Left            =   600
      TabIndex        =   3
      Top             =   3120
      Width           =   6855
      _cx             =   12091
      _cy             =   12091
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "客户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "色号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "  颜色信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Forma38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
Text1.Text = ""
Text2.Text = ""
 VSFlexGrid1.ColWidth(0) = 100
 VSFlexGrid1.ColWidth(1) = 1000
 VSFlexGrid1.ColWidth(2) = 1800
 VSFlexGrid1.ColWidth(3) = 1800
 VSFlexGrid1.BackColorAlternate = &HCDEEC6
 VSFlexGrid1.SelectionMode = flexSelectionListBox
End Sub

Private Sub VSFlexGrid1_dblClick()
    ' 捕捉错误
    On Error Resume Next
    
    ' 获取当前选中的行
    rs = VSFlexGrid1.Row
    
    ' 如果记录集为空，则退出子程序
    If Adodc1.Recordset.EOF Then Exit Sub
    
    ' 将记录集指针移到第一条记录
    Adodc1.Recordset.MoveFirst
    
    ' 将记录集指针移到选中的行
    Adodc1.Recordset.Move rs - 1
    
    ' 如果 ysbl 为 1，则设置 Formj1 的 DataCombo1
    If ysbl = 1 Then
        Formj1.DataCombo1(7).Text = Adodc1.Recordset.Fields(1)
        Formj1.DataCombo1(18).Text = Adodc1.Recordset.Fields(0)
    End If
    
    ' 如果 ysbl 为 2，则设置 Forma11 的 DataCombo4 和 Text 字段
    If ysbl = 2 Then
        Forma11.DataCombo4(6).Text = Adodc1.Recordset.Fields(2)
        Forma11.Text1 = Adodc1.Recordset.Fields(1)
        
        ' 检查 Adodc1.Recordset.Fields(2) 是否为空
        If IsNull(Adodc1.Recordset.Fields(3)) Or Adodc1.Recordset.Fields(3) = "" Then
            Forma11.Text17 = 0
        Else
            Forma11.Text17 = Adodc1.Recordset.Fields(3)
        End If
    End If
    
    ' 卸载当前窗体
    Unload Me
End Sub


Private Sub Text1_Change()
On Error Resume Next
       If Len(Text1.Text) < 3 Then Exit Sub
       Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
       Adodc1.RecordSource = "SELECT kh as 客户, sh as 色号,ys as 颜色,单价 FROM khy WHERE  sh like '%'+'" & Text1.Text & "'+'%' and kh like '%'+'" & Text2.Text & "'+'%'"
       Adodc1.Refresh
       VSFlexGrid1.ColWidth(0) = 100
       VSFlexGrid1.ColWidth(1) = 1000
       VSFlexGrid1.ColWidth(2) = 1800
       VSFlexGrid1.ColWidth(3) = 1800
       VSFlexGrid1.BackColorAlternate = &HCDEEC6
       VSFlexGrid1.SelectionMode = flexSelectionListBox
End Sub
