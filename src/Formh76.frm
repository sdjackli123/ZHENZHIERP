VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formh76 
   BackColor       =   &H00C0E0FF&
   Caption         =   "发样确认"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   8025
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   7830
      Left            =   600
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc111 
      Height          =   330
      Left            =   1920
      Top             =   8520
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
   Begin MSAdodcLib.Adodc Adodc1111 
      Height          =   330
      Left            =   1800
      Top             =   8520
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
      Left            =   5280
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   330104833
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   330104833
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   330104833
      CurrentDate     =   39177
   End
   Begin VB.Label Label1 
      Caption         =   "客户"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "日期"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "日期"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "发样日期"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Formh76"
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
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (fyr='' or fyr is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' order by sh"
Adodc1111.Refresh
Else
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (fyr='' or fyr is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH='" & Text1.Text & "' order by sh"
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
sql1 = "UPDATE KHY SET FYR='" & DTPicker3.value & "' WHERE SH='" & sehao & "'"
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
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")
Adodc1111.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (fyr='' or fyr is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH='" & Text1.Text & "' order by jyr desc "
Adodc1111.Refresh
End Sub

Private Sub Text1_Change()
t1 = Format(DTPicker1.value, "yyyy-mm-dd")
t2 = Format(DTPicker2.value, "yyyy-mm-dd")
If Text1.Text = "" Then
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (fyr='' or fyr is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' order by sh"
Adodc1111.Refresh
Else
Adodc1111.RecordSource = "SELECT ys as 颜色,sh as 色号 from khy where (fyr='' or fyr is null) AND CONVERT(varchar,jyr, 23) BETWEEN '" & t1 & "' AND '" & t2 & "' AND KH='" & Text1.Text & "' order by sh"
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

