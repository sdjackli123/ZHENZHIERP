VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Formh74 
   BackColor       =   &H00C0E0FF&
   Caption         =   "分样确认"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   7905
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Data Data1111 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   8460
      Left            =   480
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Formh74.frx":0000
      Height          =   330
      Left            =   4680
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "负责人姓名"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4680
      Top             =   8520
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
   Begin MSAdodcLib.Adodc Adodc1111 
      Height          =   330
      Left            =   4680
      Top             =   8160
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   4200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   423493633
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   423493633
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   1800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   423493633
      CurrentDate     =   39177
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "分样日期"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "日期"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "日期"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "客户"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "小样负责"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Formh74"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")
If Text1.Text = "" Then
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (dyfz='' or dyfz is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' order by sh"
Adodc1111.Refresh
Else
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (dyfz='' or dyfz is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH='" & Text1.Text & "' order by sh"
Adodc1111.Refresh
End If

If Adodc1111.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1111.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1111.Recordset.EOF
List1.AddItem Adodc1111.Recordset.Fields(1) + "/" + Trim(Adodc1111.Recordset.Fields(0))
Adodc1111.Recordset.MoveNext
Loop

End Sub

Private Sub Command3_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sehao = Mid(List1.List(i), 1, InStr(List1.List(i), "/") - 1)
sql1 = "UPDATE KHY SET dyfz='" & DataCombo1.Text & "' WHERE SH='" & sehao & "'"
RD.Open sql1, conn, adOpenStatic, adLockOptimistic
End If
Next
Call Text1_Change
End Sub

Private Sub Command8_Click()
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
DataCombo1.Text = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset

t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")
Adodc1111.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (dyfz='' or dyfz is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH='" & Text1.Text & "' order by jyr desc "
Adodc1111.Refresh
Adodc1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1.RecordSource = "select 负责人姓名 from gr group by 负责人姓名"
Adodc1.Refresh
End Sub

Private Sub Text1_Change()
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")
If Text1.Text = "" Then
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (dyfz='' or dyfz is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' order by sh"
Adodc1111.Refresh
Else
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (dyfz='' or dyfz is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH like '%'+'" & Text1.Text & "'+'%' order by sh"
Adodc1111.Refresh
End If

If Adodc1111.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1111.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1111.Recordset.EOF
List1.AddItem Adodc1111.Recordset.Fields(1) + "/" + Trim(Adodc1111.Recordset.Fields(0))
Adodc1111.Recordset.MoveNext
Loop
End Sub


