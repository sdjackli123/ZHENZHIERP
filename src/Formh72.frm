VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Formh72 
   BackColor       =   &H00C0E0FF&
   Caption         =   "发样确认"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form4"
   ScaleHeight     =   9150
   ScaleWidth      =   7305
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "刷新"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc111 
      Height          =   330
      Left            =   1560
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全选"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "全清"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "退出"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7410
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   329711619
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   329711619
      CurrentDate     =   39177
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8421376
      CalendarTrailingForeColor=   255
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss 　"
      Format          =   329711619
      CurrentDate     =   39177
   End
   Begin MSAdodcLib.Adodc Adodc1111 
      Height          =   330
      Left            =   2640
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
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "发样日期"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "日期"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "日期"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "客户"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Formh72"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection: Dim RD As ADODB.Recordset
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command3_Click()
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
sql1 = "UPDATE KHY SET FYR='" & DTPicker3.value & "' WHERE SH='" & List1.List(i) & "'"
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
Text2 = ""
Set conn = New ADODB.Connection
conn.Open "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Set RD = New ADODB.Recordset
Adodc111.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
End Sub

Private Sub Text1_Change()
Adodc1111.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
Adodc1111.RecordSource = "SELECT sh as 色号 from khy where (fyr='' or fyr is null) AND jyr BETWEEN cast('" & DTPicker1.value & "' as datetime) AND cast('" & DTPicker2.value & "' as datetime) AND KH like '%'+'" & Text1.Text & "'+'%' order by kh,jyr"
Adodc1111.Refresh
If Adodc1111.Recordset.EOF Then
List1.Clear
Exit Sub
End If
Adodc1111.Recordset.MoveFirst
List1.Clear
Do While Not Adodc1111.Recordset.EOF
List1.AddItem Adodc1111.Recordset.Fields(0)
Adodc1111.Recordset.MoveNext
Loop
End Sub
