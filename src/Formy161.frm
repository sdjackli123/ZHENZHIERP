VERSION 5.00
Begin VB.Form Formy161 
   BackColor       =   &H00C0E0FF&
   Caption         =   "注册"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   3240
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      Caption         =   "请输入注册码："
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
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Formy161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program needs 3 buttons
Const REG_none = 0
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_expand_sz = 2
Const REG_BINARY = 3 ' Free Formy binary
Const REG_dword = 4
Const REG_dword_big_endian = 5
Const REG_multi_sz = 7
Const HKEY_classes_root = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_users = &H80000003
Const HKEY_current_config = &H80000005
Const HKEY_dyn_data = &H80000006


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nFormation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
Private Sub Command1_Click()
Dim hKey As Long
Dim retu As Long


If Text1.Text = zhuce Then
RegOpenKey HKEY_LOCAL_MACHINE, "SoftWare\Microsoft\Windows\CurrentVersion\Run", hKey
retu = RegSetValueEx(hKey, "启动Word", 0, REG_SZ, ByVal Text1.Text, Len(Text1.Text))


   If retu <> 0 Then
   MsgBox ("注册未成功，不能使用本软件！")
   RegCloseKey hKey
   End
   Else
   If cpf = "614182103" And ndr = "00 00 00 00 00 00" Then
   MsgBox ("注册成功，欢迎使用本软件！")
   RegCloseKey hKey
   frmLogin.Show
   End If
   End If
Else
   MsgBox ("注册码不对，请联系软件商！")
   RegCloseKey hKey
   End
End If
RegCloseKey hKey
Unload Me
End Sub

