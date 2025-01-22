Attribute VB_Name = "ini"


Const HKEY_LOCAL_MACHINE = &H80000002

Const HKEY_users = &H80000003
Const HKEY_current_config = &H80000005

Const ERROR_SUCCESS = 0
Const REG_SZ = 1                        ' 独立的空的终结字符串
Const REG_expand_sz = 2
Const REG_BINARY = 3
Const REG_dword = 4                      ' 32位数字
Const zhuce As String = 13375564800#
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long  ''''''''调用beep
Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32
Private Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte ' Reserved, must be 0
   ncb_event      As Long
End Type
Private Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
Private Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type
Private Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long

Private Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Function GetMACAddress() As String
  'retrieve the MAC Address for the network controller
  'installed, returning a formatted string
   Dim tmp As String
   Dim pASTAT As Long
   Dim NCB As NET_CONTROL_BLOCK
   Dim AST As ASTAT
  'The IBM NetBIOS 3.0 specifications defines four basic
  'NetBIOS environments under the NCBRESET command. Win32
  'follows the OS/2 Dynamic Link Routine (DLR) environment.
  'This means that the first NCB issued by an application
  'must be a NCBRESET, with the exception of NCBENUM.
  'The Windows NT implementation differs from the IBM
  'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
   NCB.ncb_command = NCBRESET
   Call Netbios(NCB)
  'To get the Media Access Control (MAC) address for an
  'ethernet adapter programmatically, use the Netbios()
  'NCBASTAT command and provide a "*" as the name in the
  'NCB.ncb_CallName field (in a 16-chr string).
   NCB.ncb_callname = "*               "
   NCB.ncb_command = NCBASTAT
  'For machines with multiple network adapters you need to
  'enumerate the LANA numbers and perform the NCBASTAT
  'command on each. Even when you have a single network
  'adapter, it is a good idea to enumerate valid LANA numbers
  'first and perform the NCBASTAT on one of the valid LANA
  'numbers. It is considered bad programming to hardcode the
  'LANA number to 0 (see the comments section below).
   NCB.ncb_lana_num = 0
   NCB.ncb_length = Len(AST)
   pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)
   If pASTAT = 0 Then
      Debug.Print "memory allocation failed!"
      Exit Function
   End If
   NCB.ncb_buffer = pASTAT
   Call Netbios(NCB)
   CopyMemory AST, NCB.ncb_buffer, Len(AST)
'   tmp = Format$(Hex(AST.adapt.adapter_address(0)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(1)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(2)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(3)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(4)), "00") & " " & Format$(Hex(AST.adapt.adapter_address(5)), "00")
   HeapFree GetProcessHeap(), 0, pASTAT
   GetMACAddress = tmp
End Function



Sub Main()
Dim lon As Long
Dim lbj As String
Dim HKey As Long
Dim retu As Long
Dim c As String
Dim net As String

    Dim Serial As Long, VName As String, FSName As String
    'Create buffers
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    'Get the volume information
    GetVolumeInformation "C:\", VName, 255, Serial, 0, 0, FSName, 255
    'Strip the extra chr$(0)'s
    VName = Left$(VName, InStr(1, VName, Chr$(0)) - 1)
    FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
    cpf = Trim(Str$(Serial))
   'KPD-Team 2000
   'URL: http://www.allapi.net/
   'E-Mail: KPDTeam@Allapi.net
   '  -> Source: MS Knowledge Base
    ndr = GetMACAddress()


RegOpenKey HKEY_LOCAL_MACHINE, "SoftWare\Microsoft\Windows\CurrentVersion\Run", HKey
lbj = RegQueryStringValue(HKey, "启动Word")
RegCloseKey HKey
Call GetHardDiskInfo

End Sub

Function RegQueryStringValue(ByVal HKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(HKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(HKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(HKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function


