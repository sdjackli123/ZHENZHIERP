VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formd334 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Ⱦɫ����ȷ��"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   11355
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "У��"
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   3210
      ItemData        =   "Formd334.frx":0000
      Left            =   3960
      List            =   "Formd334.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ˢ��"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3840
      ItemData        =   "Formd334.frx":0004
      Left            =   600
      List            =   "Formd334.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����ȷ��"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫѡ"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫ��"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "�˳�"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "����׷��"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formd334.frx":0008
      Height          =   2415
      Left            =   600
      TabIndex        =   10
      Top             =   1320
      Width           =   6615
      _cx             =   11668
      _cy             =   4260
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
      MergeCells      =   1
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1440
      Top             =   10560
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1560
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   10560
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8280
      Top             =   10680
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   8400
      Top             =   10680
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   8640
      Top             =   10680
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
      Bindings        =   "Formd334.frx":001D
      Height          =   7455
      Left            =   7680
      TabIndex        =   11
      Top             =   1200
      Width           =   3015
      _cx             =   5318
      _cy             =   13150
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
      FormatString    =   $"Formd334.frx":0032
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ȷ��"
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
      Left            =   10560
      TabIndex        =   18
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Ⱦɫ����"
      Height          =   3375
      Left            =   3720
      TabIndex        =   17
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���������Ϣ"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "������Ϣ"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "���������"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ⱦɫ������Ϣ"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "������Ϣ"
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Formd334"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1.Recordset.EOF
List1.AddItem Trim(Adodc1.Recordset.Fields(1))
Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim L2 As String
    If MsgBox("������ѡ��ȷ�ϴ���������", vbYesNo) = vbNo Then Exit Sub

    If Text1 = "" Then Exit Sub

    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then

            ll = ""
            l1 = ""
            For Q = 0 To List3.ListCount - 1
                If List3.Selected(Q) = True Then
                    l1 = Mid(List3.List(Q), 1, InStr(List3.List(Q), "-") - 1)
                    L2 = Mid(List3.List(Q), InStr(List3.List(Q), "-") + 1) ' ȡ�� '-' ֮�������
                    bs = Val(Text3)

                    ' ��� l2 �Ƿ������ Adodc3
                    Adodc3.RecordSource = "select �������� from ghgx where ����='" & Text1.Text & "' and ���� between '1001' and '6000'"
                    Adodc3.Refresh

                    Dim exists As Boolean
                    exists = False
                    Adodc3.Recordset.MoveFirst
                    Do While Not Adodc3.Recordset.EOF
                        If Adodc3.Recordset.Fields("��������").value = L2 Then
                            exists = True
                            Exit Do
                        End If
                        Adodc3.Recordset.MoveNext
                    Loop

                    If exists Then
                        MsgBox "��ֹ�ظ��ӹ���: " & L2
                    Else
                        ' ɾ�������ƶ������ѭ���⣬ȷ������ѡ�е���ܴ���
                        'sql2 = "delete from ghgx where ����='" & Text1.Text & "' and ���='" & List1.List(i) & "' and ���� BETWEEN '1001' AND '6000'"
                        'RD.Open sql2, conn, adOpenStatic, adLockOptimistic

                        ' ִ�д洢����
                        Set g_Cmd = New Command
                        g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
                        g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
                        g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
                        g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & l1 & "','" & bs & "','" & L2 & "')"      ' ��ʾ�����ĸ��洢����
                        g_Cmd.Execute           ' ִ�д洢����
                        g_Cmd.Cancel

                        ' ����Ƿ���Ҫ���� ll
                        Adodc2.RecordSource = "select * from kpd where CHARINDEX('" & l1 & "',gx)>0"
                        Adodc2.Refresh
                        If Adodc2.Recordset.EOF Then
                            ll = ll + l1 + "-"
                        End If
                    End If
                End If
            Next

            ' ���� kpd �� gx �ֶ�
            'sql1 = "update kpd set gx=gx+'" & ll & "' where ����='" & Text1.Text & "' and ip='" & List1.List(i) & "'"
            'RD.Open sql1, conn, adOpenStatic, adLockOptimistic
        End If
    Next

    ' ���������־
    sql2 = "insert into czrz(����,����,����,����,����) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & ll & "','Ⱦɫ����')"
    RD.Open sql2, conn, adOpenStatic, adLockOptimistic

    ' ˢ��������ݿؼ�
    Adodc1.Refresh
    Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With

    MsgBox ("���óɹ���")
End Sub




Private Sub Command4_Click()
On Error Resume Next
Dim sx As Integer
If MsgBox("������ѡ��ȷ�ϴ���������", vbYesNo) = vbNo Then Exit Sub

If Text1 = "" Then Exit Sub
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then

ll = ""
l1 = ""
For Q = 0 To List3.ListCount - 1

If List3.Selected(Q) = True Then
l1 = Mid(List3.List(Q), 1, InStr(List3.List(Q), "-") - 1)
L2 = Mid(List3.List(Q), InStr(List3.List(Q), "-") + 1) ' ȡ�� '-' ֮�������
bs = Val(Text3)
Adodc3.RecordSource = "select isnull(˳��,0) from ghgx where ����='" & Text1 & "' and ���� between '1001' and '6000' order by ˳�� desc"
Adodc3.Refresh
If Adodc3.Recordset.EOF Then
sx = 1
Else
sx = Adodc3.Recordset.Fields(0) + 1
End If
Set g_Cmd = New Command
    g_Con = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    g_Cmd.ActiveConnection = g_Con          ' ���ӵ����ݿ�
    g_Cmd.CommandType = adCmdStoredProc     ' ��ʾcmd������Ϊ�洢����
    g_Cmd.CommandText = "ghgxlr('" & Text1.Text & "','" & List1.List(i) & "','" & l1 & "','" & bs & "','" & L2 & "')"     ' ��ʾ�����ĸ��洢����
    g_Cmd.Execute           ' ִ�д洢����
    g_Cmd.Cancel
    
Adodc2.RecordSource = "select * from kpd where CHARINDEX('" & l1 & "',gx)>0"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
ll = ll + l1 + "-"
End If

End If

Next

'sql1 = "update kpd set gx=gx+'" & ll & "' where ����='" & Text1.text & "' and ip='" & List1.List(i) & "'"
'RD.Open sql1, conn, adOpenStatic, adLockOptimistic

End If
Next
Adodc1.Refresh
Adodc6.Refresh

sql2 = "insert into czrz(����,����,����,����,����) VALUES('" & Now & "','" & Text1.Text & "','" & yhm & "','" & ll & "','Ⱦɫ׷��')"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

MsgBox ("���óɹ���")

End Sub

Private Sub Command5_Click()
On Error Resume Next
If Val(Text4) > 2 Then
MsgBox ("У������̫�� ��ֹ")
Exit Sub
End If
For i = 1 To VSFlexGrid2.Rows - 1
If VSFlexGrid2.Cell(flexcpChecked, i, 2) = 1 Then
bs = Val(Text4)
sql1 = "UPDATE ghgx SET ����='" & bs & "' WHERE ����='" & Text1 & "' and ����='" & VSFlexGrid2.TextMatrix(i, 2) & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Adodc6.Refresh
    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With
End Sub

Private Sub Command8_Click()
On Error Resume Next
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
End Sub

Private Sub Command9_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
End Sub

Private Sub Form_Load()
On Error Resume Next

Label2.Caption = ""
Text1.Text = ""
Text2.Text = ""
Text13.Text = ""
Text3 = 1
Text4 = 1
DataCombo1 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

If InStr(yhmk, "����") > 0 Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ����,ip as ���,Ʒ��,ɫ��,ƥ��,����,��ע,gx as ���� from kpd where ����='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"

VSFlexGrid1.ColWidth(0) = 400
VSFlexGrid1.ColWidth(8) = 3200

End Sub

Private Sub Label3_Click()
On Error Resume Next
ll = ""
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
ll = ll + Mid(List3.List(i), 1, InStr(List3.List(i), "-") - 1) + "-"
End If
Next
Text13.Text = Text13.Text + "-" + Mid(ll, 1, Len(ll) - 1)
For i = 0 To List3.ListCount - 1
List3.Selected(i) = False
Next
End Sub

Private Sub Label4_Click()
GXBL = 31
  ''''''����Ⱦ�ϵ��䷽�����Զ�����Ⱦɫ����
'FormS4.Text3 = pfyljt '''' ���������ϰѳ�̨������
If pfyl = 0 Then
FormS4.Text2 = "ˮ"
End If
If pfyl <= 0.4 And pfyl > 0 Then
FormS4.Text1 = "ǳ"
FormS4.Text2 = "��"
'FormS4.Text3 = "Ư��"
End If
If pfyl > 0.4 And pfyl <= 1.5 Then
FormS4.Text2 = "��"
End If
If pfyl > 1.5 Then
FormS4.Text2 = "��"
End If
FormS4.Show
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 4 Then
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select ����,ip as ���,Ʒ��,ɫ��,ƥ��,����,��ע,gx as ���� from kpd where ����='" & Text1.Text & "' order by ip"
Adodc1.Refresh

Adodc6.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc6.RecordSource = "select ���,����,���� from ghgx where ����='" & Text1.Text & "' and ���� between '1001' and '6000' order by ���,����"
Adodc6.Refresh

    With VSFlexGrid2
        .Editable = flexEDKbdMouse
        .Cell(flexcpChecked, 1, 2, .Rows - 1, 2) = 2
    End With

End If
Call Command2_Click
Call Command8_Click
End Sub

Private Sub Text13_Change()
List4.Clear
i = 1
For L = 0 To Int(Len(Text13.Text) / 5)
List4.AddItem Mid(Text13.Text, L * 4 + i, 4)
i = i + 1
Next
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End Sub

Private Sub Text2_Change()
Formd331.Text9 = Text2
If Text2.Text = "" Then Exit Sub
Adodc4.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc4.RecordSource = "select ���ձ��,�������� from GYSHD where ��������ϵ��='" & Text2.Text & "' and ���ձ�� between '1001' and  '6000' GROUP BY ���ձ��,�������� order by ���ձ��"
Adodc4.Refresh
If Adodc4.Recordset.EOF Then
List3.Clear
Exit Sub
End If
Adodc4.Recordset.MoveFirst
List3.Clear
Do While Not Adodc4.Recordset.EOF
List3.AddItem Adodc4.Recordset.Fields(0) + "-" + Trim(Adodc4.Recordset.Fields(1))
Adodc4.Recordset.MoveNext
Loop
For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Next
End Sub

Private Sub VSFlexGrid2_DblClick()
On Error Resume Next
If Adodc6.Recordset.EOF Then Exit Sub
rs = VSFlexGrid2.Row
Adodc6.Recordset.MoveFirst
Adodc6.Recordset.Move rs - 1
If MsgBox("ɾ��" + "��ţ�" + Trim(Adodc6.Recordset.Fields(0)) + "����" + Adodc6.Recordset.Fields(1) + "��", vbYesNo) = vbNo Then Exit Sub
sql2 = "delete from ghgx  where ����='" & Text1.Text & "' and ���='" & Adodc6.Recordset.Fields(0) & "' and ����='" & Adodc6.Recordset.Fields(1) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Adodc6.Refresh
End Sub

