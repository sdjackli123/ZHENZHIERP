VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formc141 
   BackColor       =   &H00C0E0FF&
   Caption         =   "������Ϣ"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   15960
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ѡ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3480
      Top             =   6600
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
      Left            =   4080
      Top             =   6600
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ȫ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3840
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
      Bindings        =   "Formc141.frx":0000
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   17175
      _cx             =   30295
      _cy             =   8070
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1320
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "����"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Formc141"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset: Public gygh As String

Private Sub Command1_Click()
On Error Resume Next
If Adodc1.Recordset.EOF Then Exit Sub

Adodc3.RecordSource = "SELECT ���� FROM JGMX WHERE ����='" & Text1 & "' "
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
If MsgBox("�Ѿ�������������," + Adodc3.Recordset.Fields(0) + "�Ƿ�����ظ����⣿", vbYesNo) = vbNo Then Exit Sub
End If

Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF

Adodc2.RecordSource = "SELECT ˳��� FROM JGMX WHERE ����='" & Formc15.Label13.Caption & "' ORDER BY ˳��� DESC"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
ID = Adodc2.Recordset.Fields(0) + 1
SXH = Adodc2.Recordset.Fields(0) + 1
Else
ID = 1
SXH = 1
End If

sql1 = "INSERT INTO dbo.jgmx(�ӹ���λ,Ʒ��,��ɫ,����,����,����,���,����,IP,��Լ��,˳���,����,�ӹ����,ƥ��,�ƻ���,����) Values('" & Adodc1.Recordset.Fields(0) & "','" & Adodc1.Recordset.Fields(4) & "','" & Adodc1.Recordset.Fields(5) & "','" & Adodc1.Recordset.Fields(3) & "','" & Adodc1.Recordset.Fields(6) & "','" & Adodc1.Recordset.Fields(11) & "','" & Adodc1.Recordset.Fields(12) & "','" & Date & "','" & ID & "','" & Adodc1.Recordset.Fields(2) & "','" & SXH & "','" & Formc15.Label13 & "','" & Adodc1.Recordset.Fields(8) & "','" & Adodc1.Recordset.Fields(7) & "','" & Adodc1.Recordset.Fields(1) & "','" & Formc15.DataCombo17 & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

Adodc1.Recordset.MoveNext
Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''����
sql2 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='�ѷ���' WHERE ����='" & Text1 & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("ȷ��������", vbYesNo) = vbNo Then Exit Sub
For i = 1 To VSFlexGrid1.Rows - 1
If VSFlexGrid1.Cell(flexcpChecked, i, 5) = 1 Then

Adodc3.RecordSource = "SELECT ���� FROM JGMX WHERE ����='" & Text1 & "' and �ӹ����='" & VSFlexGrid1.TextMatrix(i, 10) & "'"
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
If MsgBox("�Ѿ�������������," + Adodc3.Recordset.Fields(0) + "�Ƿ�����ظ����⣿", vbYesNo) = vbNo Then Exit Sub
End If


If Formc15.Label13.Caption = "" Then Exit Sub

Adodc2.RecordSource = "SELECT ˳��� FROM JGMX WHERE ����='" & Formc15.Label13.Caption & "' ORDER BY ˳��� DESC"
Adodc2.Refresh
If Not Adodc2.Recordset.EOF Then
ID = Adodc2.Recordset.Fields(0) + 1
SXH = Adodc2.Recordset.Fields(0) + 1
Else
ID = 1
SXH = 1
End If
                                                                                       
 
sl = Val(VSFlexGrid1.TextMatrix(i, 7))
dj = Val(VSFlexGrid1.TextMatrix(i, 12))
je = Val(VSFlexGrid1.TextMatrix(i, 13))
ps = Val(VSFlexGrid1.TextMatrix(i, 8))
'''�ͻ�����,����,isnull(��ǩ,'') as ���,����,Ʒ��,ɫ��,����,ƥ��,���,�ر�ע��,����,����,round(����*isnull(����,0),2) as ���
sql1 = "INSERT INTO dbo.jgmx(�ӹ���λ,Ʒ��,��ɫ,����,����,����,���,����,IP,��Լ��,˳���,����,�ӹ����,ƥ��,�ƻ���,����) Values('" & VSFlexGrid1.TextMatrix(i, 1) & "','" & VSFlexGrid1.TextMatrix(i, 5) & "','" & VSFlexGrid1.TextMatrix(i, 6) & "','" & VSFlexGrid1.TextMatrix(i, 4) & "','" & sl & "','" & dj & "','" & je & "','" & Date & "','" & ID & "','" & VSFlexGrid1.TextMatrix(i, 3) & "','" & SXH & "','" & Formc15.Label13.Caption & "','" & VSFlexGrid1.TextMatrix(i, 10) & "','" & ps & "','" & VSFlexGrid1.TextMatrix(i, 2) & "','" & Formc15.DataCombo17 & "')"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic

sql2 = "update dbo.kpd set FH=convert(nvarchar ,'" & Now & "',120),zt='�ѷ���' WHERE ����='" & Text1 & "' and dr='" & VSFlexGrid1.TextMatrix(i, 10) & "'"
RD.Open sql2, conn, adOpenStatic, adLockOptimistic

End If
Next

Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc3.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
VSFlexGrid1.ColWidth(0) = 100
VSFlexGrid1.ColWidth(11) = 2800
End Sub

Private Sub VSFlexGrid1_Click()
r = VSFlexGrid1.RowSel
c = VSFlexGrid1.ColSel
If c = 5 Then
If InStr(VSFlexGrid1.TextMatrix(r, 2), "Total") > 0 Then
    If VSFlexGrid1.Cell(flexcpChecked, r - Val(VSFlexGrid1.TextMatrix(r, 5)), 5, r - 1, 5) = 2 Then
            VSFlexGrid1.Cell(flexcpChecked, r - Val(VSFlexGrid1.TextMatrix(r, 5)), 5, r - 1, 5) = 1
    End If
    
End If
End If

If c = 2 Then
If InStr(VSFlexGrid1.TextMatrix(r, 2), "Total") > 0 Then
    If VSFlexGrid1.Cell(flexcpChecked, r - Val(VSFlexGrid1.TextMatrix(r, 5)), 5, r - 1, 5) = 1 Then
            VSFlexGrid1.Cell(flexcpChecked, r - Val(VSFlexGrid1.TextMatrix(r, 5)), 5, r - 1, 5) = 2
    
    End If
    
End If
End If
'If c = 2 Or c = 3 Then
'Call jc
'End If
End Sub

Private Sub VSFlexGrid1_dblClick()
On Error Resume Next
rs = VSFlexGrid1.Row
If Adodc1.Recordset.EOF Then Exit Sub
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Move rs - 1
Formc15.DataCombo1.Text = Adodc1.Recordset.Fields(0)   ''''�ͻ�
Formc15.DataCombo2.Text = Adodc1.Recordset.Fields(4)   ''Ʒ��
Formc15.DataCombo3.Text = Adodc1.Recordset.Fields(5)   '��ɫ
Formc15.DataCombo4.Text = Adodc1.Recordset.Fields(3)   '����
Formc15.DataCombo5.Text = Adodc1.Recordset.Fields(6)   'ë������
Formc15.DataCombo11.Text = Adodc1.Recordset.Fields(2)  '���
Formc15.Text7.Text = Adodc1.Recordset.Fields(7)      'ë��ƥ��
Formc15.DataCombo16.Text = Adodc1.Recordset.Fields(1) '����
Formc15.DataCombo12.Text = Adodc1.Recordset.Fields(9)    '�ر�ע��
Formc15.Combo1.Text = Adodc1.Recordset.Fields(8)  '���
'Formc15.Text9.Text = Adodc1.Recordset.Fields(15)  Ӣ��ɫ��
Formc15.Text8.Text = Adodc1.Recordset.Fields(11)   ''''����
Unload Me
End Sub

Private Sub Text1_Change()
On Error Resume Next
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select �ͻ�����,����,isnull(��ǩ,'') as ���,����,Ʒ��,ɫ��,����,ƥ��,���,�ر�ע��,����,����,round(����*isnull(����,0),2) as ���,�Ƿ񷢻� from v_kpd_fh  WHERE ����='" & Text1.Text & "'"
Adodc1.Refresh
VSFlexGrid1.AutoSizeMode = flexAutoSizeRowHeight
VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1, False, 30

If Adodc1.Recordset.EOF Then
hs = 0
Else
hs = Adodc1.Recordset.RecordCount + 1
End If

If hs > 0 Then
    With VSFlexGrid1
        .Editable = flexEDKbdMouse
'        .AutoSize 0
        .Cell(flexcpChecked, 1, 5, hs - 1, 5) = 2
'        .Cell(MergeCells, 1, 2, hs - 1, 2) = True
        End With
VSFlexGrid1.SubtotalPosition = flexSTBelow
VSFlexGrid1.Subtotal flexSTSum, 2, 7, , vbGreen
VSFlexGrid1.Subtotal flexSTSum, 2, 8, , vbGreen
VSFlexGrid1.Subtotal flexSTCount, 2, 5, , vbGreen
End If

End Sub

Private Sub jc()
sl1 = 0
sl2 = 0
For i = 1 To VSFlexGrid1.Rows - 1
If VSFlexGrid1.Cell(flexcpChecked, i, 3) = 1 Then
sl1 = sl1 + 1
sl2 = sl2 + Val(VSFlexGrid1.TextMatrix(i, 4))
End If
Next
End Sub
