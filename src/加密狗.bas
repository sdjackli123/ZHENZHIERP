Attribute VB_Name = "加密狗"
Option Explicit


Public DISKNO As String
Private Const OFFSET_4 = 4294967296#    '&h100000000 的数值
Private Const MAXINT_4 = 2147483647     '整型数据所能表示的最大正数值 (&h7fffffff)
Private Const Bit_32 = 2147483648#      '&h80000000 的正整数形式
Private Const DELTA = &H9E3779B9        'TEA算法的DELTA值
Private Declare Sub CopyByteArray Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As Long, ByVal ByteLen As Long)
Private Declare Sub CopyStringToByte Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As String, ByVal ByteLen As Long)
Private Declare Sub CopyByteToString Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As String, ByVal pSource As Long, ByVal ByteLen As Long)

Private Type GUID
    Data1(1 To 4) As Byte
    Data2(1 To 2) As Byte
    Data3(1 To 2) As Byte
    Data4(1 To 8) As Byte
End Type

Private Type SP_INTERFACE_DEVICE_DATA
     cbSize As Long
     InterfaceClassGuid As GUID
     Flags As Long
     Reserved As Long
End Type

Private Type SP_DEVINFO_DATA
     cbSize As Long
     ClassGuid As GUID
     DevInst As Long
     Reserved As Long
End Type

Private Type SP_DEVICE_INTERFACE_DETAIL_DATA
   cbSize As Long
   DevicePath(0 To 255) As Byte
End Type

Private Type HIDD_ATTRIBUTES
    Size As Long
      VendorID As Integer
      ProductID As Integer
      VersionNumber As Integer
End Type
 
Private Type HIDP_CAPS
       Usage As Integer
       UsagePage As Integer
       InputReportByteLength As Integer
       OutputReportByteLength As Integer
       FeatureReportByteLength As Integer
       Reserved(1 To 17) As Integer

       NumberLinkCollectionNodes As Integer

       NumberInputButtonCaps As Integer
       NumberInputValueCaps As Integer
       NumberInputDataIndices As Integer

       NumberOutputButtonCaps As Integer
       NumberOutputValueCaps As Integer
       NumberOutputDataIndices As Integer

       NumberFeatureButtonCaps As Integer
       NumberFeatureValueCaps As Integer
       NumberFeatureDataIndices As Integer
End Type
 
Private Const VID = &H3689
Private Const PID = &H8762
Private Const DIGCF_PRESENT = &H2
Private Const DIGCF_DEVICEINTERFACE = &H10
Private Const INVALID_HANDLE_VALUE = (-1)
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INFINITE = &HFFFF      '  Infinite timeout

Private Const MAX_LEN = 495


Private Declare Function HidD_GetAttributes Lib "HID.dll" (ByVal HidDeviceObject As Long, ByRef Attributes As HIDD_ATTRIBUTES) As Boolean

Private Declare Function HidD_GetHidGuid Lib "HID.dll" (ByRef HidGuid As GUID) As Long

Private Declare Function SetupDiGetClassDevsA Lib "SetupApi.dll" (ByRef ClassGuid As GUID, ByVal Enumerator As Long, ByVal hwndParent As Long, ByVal Flags As Long) As Long

Private Declare Function SetupDiDestroyDeviceInfoList Lib "SetupApi.dll" (ByVal DeviceInfoSet As Long) As Boolean

Private Declare Function SetupDiGetDeviceInterfaceDetailA Lib "SetupApi.dll" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_INTERFACE_DEVICE_DATA, ByRef DeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA, _
                                                    ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Boolean

Private Declare Function SetupDiGetDeviceInterfaceDetail_2 Lib "SetupApi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_INTERFACE_DEVICE_DATA, ByVal DeviceInterfaceDetailData As Long, _
                                                    ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Boolean


Private Declare Function SetupDiEnumDeviceInterfaces Lib "SetupApi.dll" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_INTERFACE_DEVICE_DATA) As Boolean

Private Declare Function HidD_GetPreparsedData Lib "HID.dll" (ByVal HidDeviceObject As Long, ByRef PreparsedData As Long) As Boolean

Private Declare Function HidP_GetCaps Lib "HID.dll" (ByVal PreparsedData As Long, ByRef Capabilities As HIDP_CAPS) As Long

Private Declare Function HidD_FreePreparsedData Lib "HID.dll" (ByVal PreparsedData As Long) As Boolean

Private Declare Function HidD_SetFeature Lib "HID.dll" (ByVal HidDeviceObject As Long, ByVal ReportBuffer As Long, ByVal ReportBufferLength As Long) As Boolean
   
Private Declare Function HidD_GetFeature Lib "HID.dll" (ByVal HidDeviceObject As Long, ByVal ReportBuffer As Long, ByVal ReportBufferLength As Long) As Boolean

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (ByVal lpSemaphoreAttributes As Long, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function ReleaseSemaphore Lib "kernel32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal InString As String) As Long

Dim CF As Long
Private Function myhex(indata As Byte) As String
Dim s As String
s = Hex(indata)
If Len(s) < 2 Then s = "0" + s
myhex = s
End Function
Public Function StringFromBuffer(Buffer As String) As String
    Dim nPos As Long

    nPos = InStr(Buffer, vbNullChar)
    If nPos > 0 Then
        StringFromBuffer = Left$(Buffer, nPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function
Private Function AddLong(lX As Long, lY As Long) As Long '长整数加法函数
    Dim lX4 As Long
    Dim lY4 As Long
    Dim lX8 As Long
    Dim lY8 As Long
    Dim lResult As Long
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddLong = lResult
End Function

Private Function SubtractLong(lX As Long, lY As Long) As Long '长整数减法函数
    Dim lX8 As Long
    Dim lY8 As Long
    Dim mX As Double
    Dim mY As Double
    Dim mResult As Double
    Dim lResult As Long
    
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    
    mX = lX And &H7FFFFFFF
    mY = lY And &H7FFFFFFF
    
    If lX8 Then
       If lY8 Then
          mResult = mX - mY
       Else
          mX = mX + Bit_32
          mResult = mX - mY
       End If
    Else
       If lY8 Then
          mY = lY
          mResult = mX - mY
       Else
          mResult = mX - mY
       End If
    End If
    
    
    If mResult < 0 Then
       lResult = ((Bit_32 + mResult) Or &H80000000) And &HFFFFFFFF
    ElseIf mResult > MAXINT_4 Then
       lResult = ((mResult - Bit_32) Or &H80000000) And &HFFFFFFFF
    Else
       lResult = mResult And &HFFFFFFFF
    End If
    
    SubtractLong = lResult
 
End Function

Private Function LeftRotateLong(ByVal lValue As Long, lBits As Integer) As Long '按位左移函数
    Dim lngSign As Long, intI As Integer
    Dim mValue As Long
    
    lBits = lBits Mod 32
    mValue = lValue
    If lBits = 0 Then LeftRotateLong = mValue: Exit Function
    
    For intI = 1 To lBits
        lngSign = mValue And &H40000000
        mValue = (mValue And &H3FFFFFFF) * 2
     
        If lngSign And &H40000000 Then
           mValue = mValue Or &H80000000
        End If
    Next
    
    LeftRotateLong = mValue
End Function


Private Function RightRotateLong(ByVal lValue As Long, lBits As Integer) As Long '按位右移函数
   Dim lngSign As Long, intI As Integer
   Dim mValue As Long
   
   mValue = lValue
   lBits = lBits Mod 32
   
   If lBits = 0 Then RightRotateLong = mValue: Exit Function
   
   For intI = 1 To lBits
      lngSign = mValue And &H80000000
      mValue = (mValue And &H7FFFFFFF) \ 2
      If lngSign Then
         mValue = mValue Or &H40000000
      End If
   Next
   RightRotateLong = mValue
End Function
Public Sub Encode(b() As Byte, outb() As Byte, Key As String) '增强算法--加密
Dim KeyBuf(16) As Byte
Dim v(2) As Long
Dim k(4) As Long
HexStringToLongArray Key, k
ByteToLong v, b
sub_Encode v, k
LongToByte v, outb
End Sub
Public Sub Decode(b() As Byte, outb() As Byte, Key As String) '增强算法--解密
Dim KeyBuf(16) As Byte
Dim v(2) As Long
Dim k(4) As Long
HexStringToLongArray Key, k
ByteToLong v, b
sub_Decode v, k
LongToByte v, outb
End Sub
Private Sub sub_Encode(v() As Long, k() As Long)
   Dim y As Long, Z As Long
   Dim K1 As Long, K2 As Long, K3 As Long, K4 As Long
   Dim l1 As Long, L2 As Long, L3 As Long, L4 As Long
   
   Dim Sum As Long
   Dim i As Integer, Rounds As Integer
   Dim mResult(0 To 1) As Long
      
   y = v(0)
   Z = v(1)
   K1 = k(0)
   K2 = k(1)
   K3 = k(2)
   K4 = k(3)
   

  Rounds = 32
   
   For i = 1 To Rounds
      'sum += delta ;
      Sum = AddLong(Sum, DELTA)
      'y += (z<<4)+k[0] ^ z+sum ^ (z>>5)+k[1]
      l1 = LeftRotateLong(Z, 4)
      l1 = AddLong(l1, K1)
      L2 = AddLong(Z, Sum)
      L3 = RightRotateLong(Z, 5)
      L3 = AddLong(L3, K2)
      L4 = l1 Xor L2
      L4 = L4 Xor L3
      y = AddLong(y, L4)
      'z += (y<<4)+k[2] ^ y+sum ^ (y>>5)+k[3]
      l1 = LeftRotateLong(y, 4)
      l1 = AddLong(l1, K3)
      L2 = AddLong(y, Sum)
      L3 = RightRotateLong(y, 5)
      L3 = AddLong(L3, K4)
      L4 = l1 Xor L2 Xor L3
      Z = AddLong(Z, L4)
   Next
   
   v(0) = y
   v(1) = Z
End Sub

Private Sub sub_Decode(v() As Long, k() As Long) '增强算法--解密
   Dim y As Long, Z As Long
   Dim K1 As Long, K2 As Long, K3 As Long, K4 As Long
   Dim l1 As Long, L2 As Long, L3 As Long, L4 As Long
   Dim Sum As Long
   Dim i As Integer, Rounds As Integer
   Dim mResult(0 To 1) As Long
      
   y = v(0)
   Z = v(1)
   K1 = k(0)
   K2 = k(1)
   K3 = k(2)
   K4 = k(3)
   

   Rounds = 32
  Sum = LeftRotateLong(DELTA, 5)
   
   For i = 1 To Rounds

      l1 = LeftRotateLong(y, 4)
      l1 = AddLong(l1, K3)
      L2 = AddLong(y, Sum)
      L3 = RightRotateLong(y, 5)
      L3 = AddLong(L3, K4)
      L4 = l1 Xor L2 Xor L3
      Z = SubtractLong(Z, L4)
      
      l1 = LeftRotateLong(Z, 4)
      l1 = AddLong(l1, K1)
      L2 = AddLong(Z, Sum)
      L3 = RightRotateLong(Z, 5)
      L3 = AddLong(L3, K2)
      L4 = l1 Xor L2 Xor L3
      y = SubtractLong(y, L4)

      Sum = SubtractLong(Sum, DELTA)
   
   Next

  v(0) = y
  v(1) = Z
End Sub
Private Sub ByteToLong(v() As Long, b() As Byte)
Dim n As Integer
v(0) = 0
v(1) = 0
For n = 0 To 3
   v(0) = yt_movebitleft(b(n), n * 8) Or v(0)
   v(1) = yt_movebitleft(b(n + 4), n * 8) Or v(1)
Next n
End Sub

Private Sub LongToByte(v() As Long, b() As Byte)
Dim n As Integer
Dim temp As Byte
For n = 0 To 3
  b(n) = yt_movebitright(v(0), n * 8) And 255
  b(n + 4) = yt_movebitright(v(1), n * 8) And 255
Next n

End Sub
Private Sub HexStringToLongArray(ByVal Key As String, k() As Long)
Dim nlen As Integer
Dim n As Integer
Dim temp As String
Dim buf(16) As Long
nlen = Len(Key)
Dim i As Integer
i = 0
For n = 1 To nlen Step 2
   temp = Mid(Key, n, 2)
   buf(i) = HexToInt(temp)
   i = i + 1
Next n
For n = 0 To 3
   k(n) = 0
Next n
For n = 0 To 3
   k(0) = yt_movebitleft(buf(n), n * 8) Or k(0)
   k(1) = yt_movebitleft(buf(n + 4), n * 8) Or k(1)
   k(2) = yt_movebitleft(buf(n + 4 + 4), n * 8) Or k(2)
   k(3) = yt_movebitleft(buf(n + 4 + 4 + 4), n * 8) Or k(3)
Next n

End Sub


Private Function ByteArrayToHexString(b() As Byte, nlen As Long) As String
Dim outstring As String
outstring = ""
Dim n As Integer
Dim temp As Byte
For n = 0 To nlen - 1
  outstring = outstring + myhex(b(n))
Next n
ByteArrayToHexString = outstring
End Function

Private Function LongArrayToHexString(k() As Long) As String
Dim outstring As String
outstring = ""
Dim nlen As Integer
Dim n As Integer
Dim temp As Byte
For n = 0 To 3
  temp = yt_movebitleft(k(0), n * 8)
  outstring = outstring + myhex(temp)
Next n
For n = 0 To 3
   temp = yt_movebitleft(k(1), n * 8)
   outstring = outstring + myhex(temp)
Next n
LongArrayToHexString = outstring
End Function

Private Function HexToInt(ByVal s As String) As Long
Dim hexch As String
Dim temp As String
Dim a As Long
Dim i As Long
Dim J As Long
Dim r As Long
Dim n As Long
Dim k As Long
s = UCase(s)
hexch = "0123456789ABCDEF"
k = 1
r = 0
For i = Len(s) To 1 Step -1
    temp = Mid(s, i, 1)
    n = 0
        For J = 1 To 16
        If temp = Mid(hexch, J, 1) Then n = J - 1
        Next J
        r = r + (n * k)
        k = k * 16
        
Next i
    HexToInt = r
End Function


'进位循环左移

Private Function RCL(OPR As Byte) As Byte
Dim BD As Byte
Dim i As Integer
Dim Fg1 As Byte
Dim Fg2 As Byte
BD = OPR
Fg2 = CF And 1
Fg1 = (BD And &H80) \ 128 '判断D7位是否进位
BD = ((BD And &H7F) * 2) Or Fg2 '带进位左移
Fg2 = Fg1
CF = Fg1
RCL = BD
End Function
 

'进位循环右移

Private Function RCR(OPR As Byte) As Byte
Dim BD As Byte
Dim Fg1 As Byte
Dim Fg2 As Byte
BD = OPR
Fg2 = CF And 128
Fg1 = (BD And 1) * 128 '判断D0位是否进位
BD = (BD \ 2) Or Fg2 '带进位右移
Fg2 = Fg1
CF = Fg1
RCR = BD
End Function
 
'循环左移
Private Function yt_movebitleft(ByVal OPR As Long, movebit As Integer) As Long
Dim temp As Long
temp = OPR
Dim i As Integer
Dim m As Integer
Dim ByteArray(3) As Byte
CopyByteArray VarPtr(ByteArray(0)), VarPtr(temp), 4
For m = 1 To movebit
    CF = 0
    RCL ByteArray(3)
    For i = 0 To 3
    ByteArray(i) = RCL(ByteArray(i))
    Next i
Next m
CopyByteArray VarPtr(temp), VarPtr(ByteArray(0)), 4
yt_movebitleft = temp
End Function

'循环右移
Private Function yt_movebitright(ByVal OPR As Long, movebit As Integer) As Long
Dim temp As Long
temp = OPR
Dim i As Integer
Dim m As Integer
Dim ByteArray(3) As Byte
CopyByteArray VarPtr(ByteArray(0)), VarPtr(temp), 4
For m = 1 To movebit
    CF = 0
    RCR ByteArray(0)
    For i = 3 To 0 Step -1
        ByteArray(i) = RCR(ByteArray(i))
    Next i
Next m
CopyByteArray VarPtr(temp), VarPtr(ByteArray(0)), 4
yt_movebitright = temp
End Function
Public Function StrDec(InString As String, Key As String) As String '使用增强算法，解密字符串
Dim b() As Byte
Dim outb() As Byte
Dim temp(8) As Byte
Dim outtemp(8) As Byte
Dim n As Long
Dim nlen As Long
nlen = HexStringToByteArray(InString, b)
ReDim outb(nlen)

For n = 0 To nlen
    outb(n) = b(n)
Next n

For n = 0 To nlen - 8 Step 8
    MoveByte temp, b, n
    Decode temp, outtemp, Key
    MoveByte_2 outb, outtemp, n
Next n
Dim outstring As String
outstring = Space(nlen + 1)
CopyByteToString outstring, VarPtr(outb(0)), nlen
StrDec = StringFromBuffer(outstring)
End Function

Public Function StrEnc(InString As String, Key As String) As String '使用增强算法，加密字符串
Dim b() As Byte
Dim outb() As Byte
Dim temp(8) As Byte
Dim outtemp(8) As Byte
Dim n As Long
Dim nlen As Long
nlen = StringToByte(b, InString)
ReDim outb(nlen)

For n = 0 To nlen
    outb(n) = b(n)
Next n

For n = 0 To nlen - 8 Step 8
    MoveByte temp, b, n
    Encode temp, outtemp, Key
    MoveByte_2 outb, outtemp, n
Next n
StrEnc = ByteArrayToHexString(outb, nlen)
End Function
Private Sub MoveByte_2(destb() As Byte, orgb() As Byte, pos As Long)
Dim n As Long
For n = 0 To 7
    destb(n + pos) = orgb(n)
Next n
End Sub

Private Sub MoveByte(destb() As Byte, orgb() As Byte, pos As Long)
Dim n As Long
For n = 0 To 7
    destb(n) = orgb(n + pos)
Next n
End Sub

Private Function StringToByte(b() As Byte, ByVal InString As String) As Long
Dim nlen As Long
Dim n As Long
nlen = lstrlen(InString) + 1
If nlen < 8 Then
    StringToByte = 8
Else
    StringToByte = nlen
 End If
ReDim b(StringToByte)
CopyStringToByte VarPtr(b(0)), InString, nlen
End Function

Private Function HexStringToByteArray(ByVal InString As String, b() As Byte) As Long
Dim nlen As Integer
Dim n As Integer
Dim temp As String
nlen = Len(InString)
If nlen < 16 Then HexStringToByteArray = 16
HexStringToByteArray = nlen / 2
ReDim b(HexStringToByteArray)
Dim i As Integer
i = 0
For n = 1 To nlen Step 2
   temp = Mid(InString, n, 2)
   b(i) = HexToInt(temp)
   i = i + 1
Next n
End Function

Private Function GetFeature(ByVal hDevice As Long, array_out() As Byte, ByVal out_len As Integer) As Boolean

    Dim FeatureStatus As Boolean
    Dim Status As Boolean
    Dim i As Integer
    Dim FeatureReportBuffer(0 To 512) As Byte
    Dim Ppd As Long
    Dim Caps As HIDP_CAPS
    
    If HidD_GetPreparsedData(hDevice, Ppd) = False Then Exit Function

    If HidP_GetCaps(Ppd, Caps) = False Then
            HidD_FreePreparsedData Ppd
            Exit Function
    End If

    Status = True

    FeatureReportBuffer(0) = 1

    FeatureStatus = HidD_GetFeature(hDevice, VarPtr(FeatureReportBuffer(0)), Caps.FeatureReportByteLength)
    If FeatureStatus Then
            For i = 0 To out_len - 1
                array_out(i) = FeatureReportBuffer(i)
            Next i
    End If

    Status = Status And FeatureStatus
    HidD_FreePreparsedData Ppd

   GetFeature = Status
End Function

Private Function SetFeature(ByVal hDevice As Long, array_in() As Byte, ByVal in_len As Integer) As Boolean
    Dim FeatureStatus As Boolean
    Dim Status As Boolean
    Dim i As Integer
     Dim FeatureReportBuffer(0 To 512) As Byte
    Dim Ppd As Long
    Dim Caps As HIDP_CAPS
    Dim ret As Long

     If HidD_GetPreparsedData(hDevice, Ppd) = False Then Exit Function
     
    If HidP_GetCaps(Ppd, Caps) = False Then
            HidD_FreePreparsedData Ppd
            Exit Function
    End If

    Status = True

    FeatureReportBuffer(0) = 2
  
    For i = 0 To in_len - 1
        FeatureReportBuffer(i + 1) = array_in(i + 1)

    Next i
    FeatureStatus = HidD_SetFeature(hDevice, VarPtr(FeatureReportBuffer(0)), Caps.FeatureReportByteLength)


    Status = Status And FeatureStatus
    HidD_FreePreparsedData Ppd

   SetFeature = Status
End Function

Private Function isfindmydevice(ByVal pos As Integer, ByRef count As Integer, ByRef OutPath As String) As Boolean
    Dim hardwareDeviceInfo As Long
    Dim DeviceInfoData As SP_INTERFACE_DEVICE_DATA
    Dim i As Long
    Dim HidGuid As GUID
    Dim functionClassDeviceData As SP_DEVICE_INTERFACE_DETAIL_DATA
    Dim requiredLength                                As Long
    Dim d_handle As Long
    Dim Attributes As HIDD_ATTRIBUTES

    HidD_GetHidGuid HidGuid

    hardwareDeviceInfo = SetupDiGetClassDevsA(HidGuid, 0, 0, DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE)

    If (hardwareDeviceInfo = INVALID_HANDLE_VALUE) Then Exit Function

    DeviceInfoData.cbSize = Len(DeviceInfoData)

    Do While (SetupDiEnumDeviceInterfaces(hardwareDeviceInfo, 0, HidGuid, i, DeviceInfoData))
            If GetLastError = ERROR_NO_MORE_ITEMS Then Exit Do
                functionClassDeviceData.cbSize = 5
            If SetupDiGetDeviceInterfaceDetailA(hardwareDeviceInfo, DeviceInfoData, functionClassDeviceData, 300, requiredLength, 0) = False Then
                SetupDiDestroyDeviceInfoList hardwareDeviceInfo
                Exit Function
            End If
            OutPath = Space(260)
            lstrcpy OutPath, VarPtr(functionClassDeviceData.DevicePath(0))
            d_handle = CreateFile(OutPath, 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
            If INVALID_HANDLE_VALUE <> d_handle Then
                 If HidD_GetAttributes(d_handle, Attributes) Then
                        If (Attributes.ProductID = PID) And (Attributes.VendorID = VID) Then
                                If pos = count Then
                                    SetupDiDestroyDeviceInfoList hardwareDeviceInfo
                                     CloseHandle d_handle
                                    isfindmydevice = True: Exit Function
                                End If
                                count = count + 1
                        End If
                End If
                 CloseHandle d_handle
            End If
            i = i + 1
       Loop
    SetupDiDestroyDeviceInfoList hardwareDeviceInfo
End Function

Private Function NT_FindPort(ByVal start As Integer, ByRef OutPath As String) As Integer
    Dim count As Integer
    If Not isfindmydevice(start, count, OutPath) Then
        NT_FindPort = -92
    End If
End Function


Private Function OpenMydivece(ByRef hUsbDevice As Long, ByVal Path As String) As Integer
    Dim OutPath As String
    Dim biao As Boolean
    Dim count As Integer
    If Len(Path) < 1 Then
        biao = isfindmydevice(0, count, OutPath)
        If biao = False Then OpenMydivece = -92: Exit Function
        hUsbDevice = CreateFile(OutPath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If hUsbDevice = INVALID_HANDLE_VALUE Then OpenMydivece = -92: Exit Function
    Else
        hUsbDevice = CreateFile(Path, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If hUsbDevice = INVALID_HANDLE_VALUE Then OpenMydivece = -92: Exit Function
    End If
End Function

 Private Function GetIDVersion(ByRef Version As Integer, ByVal Path As String) As Integer

    Dim ret As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
    array_in(1) = 1
    If SetFeature(hUsbDevice, array_in, 1) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    If GetFeature(hUsbDevice, array_out, 1) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    CloseHandle (hUsbDevice)
    Version = array_out(0)
    Exit Function
error_exit:
    GetIDVersion = ret
End Function

 Private Function NT_GetID(ByRef ID_1 As Long, ByRef ID_2 As Long, ByVal Path As String) As Integer
    Dim ret As Integer
    Dim t(0 To 8) As Long
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
    array_in(1) = 2
    If SetFeature(hUsbDevice, array_in, 1) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    If GetFeature(hUsbDevice, array_out, 8) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    CloseHandle (hUsbDevice)
    t(0) = array_out(0): t(1) = array_out(1): t(2) = array_out(2): t(3) = array_out(3)
    t(4) = array_out(4): t(5) = array_out(5): t(6) = array_out(6): t(7) = array_out(7)
    ID_1 = t(3) Or LeftRotateLong(t(2), 8) Or LeftRotateLong(t(1), 16) Or LeftRotateLong(t(0), 24)
    ID_2 = t(7) Or LeftRotateLong(t(6), 8) Or LeftRotateLong(t(5), 16) Or LeftRotateLong(t(4), 24)
    Exit Function
error_exit:
    NT_GetID = ret
End Function


 Private Function Y_Read(ByRef OutData() As Byte, ByVal address As Integer, ByVal nlen As Integer, ByRef Password() As Byte, ByVal Path As String, ByVal pos As Integer) As Integer

    Dim addr_l As Integer
    Dim addr_h As Integer
    Dim n As Integer
    Dim ret As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    If (address > MAX_LEN) Or (address < 0) Then ret = -81: GoTo error_exit
    If (nlen > 16) Then ret = -87: GoTo error_exit
    If (nlen + address) > MAX_LEN Then ret = -88: GoTo error_exit
    addr_h = RightRotateLong(address, 8) * 2
    addr_l = address And 255
    Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit

        array_in(1) = &H12
        array_in(2) = addr_h
        array_in(3) = addr_l
        array_in(4) = nlen
        For n = 0 To 7
            array_in(5 + n) = Password(n)
        Next n
    If SetFeature(hUsbDevice, array_in, 13) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    If GetFeature(hUsbDevice, array_out, nlen + 1) = False Then CloseHandle (hUsbDevice): ret = -94: GoTo error_exit
    CloseHandle (hUsbDevice)
    If array_out(0) <> 0 Then
           Y_Read = -83: Exit Function
    End If
    For n = 0 To (nlen - 1)
        OutData(n + pos) = array_out(n + 1)
    Next n
    Exit Function
error_exit:
    Y_Read = ret
End Function

 Private Function Y_Write(ByRef indata() As Byte, ByVal address As Integer, ByVal nlen As Integer, ByRef Password() As Byte, ByVal Path As String, ByVal pos As Integer) As Integer
    Dim ret As Integer
    Dim addr_l As Integer
    Dim addr_h As Integer
    Dim n As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    If (nlen > 8) Then ret = -87: GoTo error_exit
    If ((address + nlen - 1) > (MAX_LEN + 17)) Or (address < 0) Then ret = -81: GoTo error_exit
    addr_h = RightRotateLong(address, 8) * 2
    addr_l = address And 255
    Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
       array_in(1) = &H13
        array_in(2) = addr_h
        array_in(3) = addr_l
        array_in(4) = nlen
        For n = 0 To 7
            array_in(5 + n) = Password(n)
        Next n
        For n = 0 To nlen - 1
            array_in(13 + n) = indata(n + pos)
        Next n
    If SetFeature(hUsbDevice, array_in, 13 + nlen) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    If GetFeature(hUsbDevice, array_out, 2) = False Then CloseHandle (hUsbDevice): ret = -94: GoTo error_exit
    CloseHandle (hUsbDevice)
    If array_out(0) <> 0 Then
           Y_Write = -82
    End If
    Exit Function
error_exit:
    Y_Write = ret
End Function

 Private Function NT_Cal(ByRef InBuf() As Byte, ByRef OutBuf() As Byte, ByVal Path As String, ByVal pos As Integer) As Integer
    Dim ret As Integer
    Dim n As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
    array_in(1) = 8
    For n = 2 To 9
        array_in(n) = InBuf(n - 2 + pos)
    Next n
    If SetFeature(hUsbDevice, array_in, 9) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    If GetFeature(hUsbDevice, array_out, 9) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    CloseHandle (hUsbDevice)
    For n = 0 To 7
       OutBuf(n + pos) = array_out(n)
    Next n
    If array_out(8) <> &H55 Then
        NT_Cal = -20
    End If
        Exit Function
error_exit:
    NT_Cal = ret
End Function

 Private Function NT_SetCal_2(ByRef indata() As Byte, ByVal IsHi As Byte, ByVal Path As String, ByVal pos As Integer) As Integer
       Dim ret As Integer
       Dim n As Integer
       Dim array_in(0 To 30) As Byte
       Dim array_out(0 To 25) As Byte
       Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
       array_in(1) = 9
       array_in(2) = IsHi
       For n = 0 To 7
           array_in(3 + n) = indata(n + pos)
       Next n
       If SetFeature(hUsbDevice, array_in, 11) = False Then CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
       If GetFeature(hUsbDevice, array_out, 2) = False Then CloseHandle (hUsbDevice): ret = -94: GoTo error_exit
       CloseHandle (hUsbDevice)
       If array_out(0) <> 0 Then
           NT_SetCal_2 = -82
       End If
       
    Exit Function
error_exit:
    NT_SetCal_2 = ret
End Function

Public Function NT_GetIDVersion(ByRef Version As Integer, ByVal Path As String) As Integer
    Dim hsignal As Long
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    NT_GetIDVersion = GetIDVersion(Version, Path)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function GetID(ByRef ID_1 As Long, ByRef ID_2 As Long, ByVal Path As String) As Integer
    Dim hsignal As Long
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    GetID = NT_GetID(ID_1, ID_2, Path)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function YWriteEx(ByRef indata() As Byte, ByVal address As Integer, ByVal nlen As Integer, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim Password(0 To 8) As Byte
    Dim n As Integer
    Dim leave As Integer
    Dim temp_leave As Integer
    If (address + nlen - 1 > MAX_LEN) Or (address < 0) Then YWriteEx = -81: Exit Function
    myconvert HKey, LKey, Password
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    temp_leave = address Mod 16
    leave = 16 - temp_leave
    If leave > nlen Then leave = nlen
    If (leave > 0) Then
        For n = 0 To leave \ 8 - 1
            YWriteEx = Y_Write(indata, address + n * 8, 8, Password, Path, 8 * n)
            If (YWriteEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        Next n
        If (leave - 8 * n > 0) Then
            YWriteEx = Y_Write(indata, address + n * 8, (leave - n * 8), Password, Path, 8 * n)
            If (YWriteEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        End If
    End If
    nlen = nlen - leave: address = address + leave
    If (nlen > 0) Then
        For n = 0 To nlen \ 8 - 1
            YWriteEx = Y_Write(indata, address + n * 8, 8, Password, Path, leave + 8 * n)
            If (YWriteEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        Next n
        If (nlen - 8 * n > 0) Then
            YWriteEx = Y_Write(indata, address + n * 8, (nlen - n * 8), Password, Path, leave + 8 * n)
            If (YWriteEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        End If
    End If
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function
            
Public Function YReadEx(ByRef OutData() As Byte, ByVal address As Integer, ByVal nlen As Integer, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim Password(0 To 8) As Byte
    Dim n As Integer
    Dim i As Integer
    Dim outlen As Integer
    If (address + nlen - 1 > MAX_LEN) Or (address < 0) Then YReadEx = -81: Exit Function
    myconvert HKey, LKey, Password
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    For n = 0 To nlen \ 16 - 1
        YReadEx = Y_Read(OutData, address + n * 16, 16, Password, Path, n * 16)
        If (YReadEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
    Next n
    If (nlen - 16 * n > 0) Then
        YReadEx = Y_Read(OutData, address + n * 16, (nlen - 16 * n), Password, Path, 16 * n)
        If (YReadEx <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
    End If
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function



Public Function FindPort(ByVal start As Integer, ByRef OutPath As String) As Integer
    Dim hsignal As Long
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    FindPort = NT_FindPort(start, OutPath)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function


Private Function AddZero(ByVal InKey As String) As String
Dim nlen As Integer
Dim n As Integer
nlen = Len(InKey)
For n = nlen To 7
  InKey = "0" + InKey
Next n
AddZero = InKey
End Function

Private Sub myconvert(ByVal HKey As String, ByVal LKey As String, ByRef out_data() As Byte)
HKey = AddZero(HKey)
LKey = AddZero(LKey)
  Dim n As Integer
    For n = 0 To 3
        out_data(n) = HexToInt(Mid(HKey, 1 + n * 2, 2))
    Next n
    For n = 0 To 3
        out_data(n + 4) = HexToInt(Mid(LKey, 1 + n * 2, 2))
    Next n
End Sub
Public Function YRead(ByRef indata As Byte, ByVal address As Integer, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim ary1(0 To 8) As Byte

    If (address > 495) Or (address < 0) Then YRead = -81: Exit Function
    myconvert HKey, LKey, ary1
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    YRead = sub_YRead(indata, address, ary1, Path)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function
Private Function sub_YRead(ByRef OutData As Byte, ByVal address As Integer, ByRef Password() As Byte, ByVal Path As String) As Integer
    Dim addr_l As Integer
    Dim addr_h As Integer
    Dim n As Integer
    Dim ret As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    Dim hUsbDevice As Long
    Dim opcode As Byte
    If (address > 495) Or (address < 0) Then ret = -81: GoTo error_exit
    opcode = 128
    If (address > 255) Then
     opcode = 160
     address = address - 256
    End If
    
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
        array_in(1) = 16
        array_in(2) = opcode
        array_in(3) = address
        array_in(4) = address
        For n = 0 To 7
            array_in(5 + n) = Password(n)
        Next
    If SetFeature(hUsbDevice, array_in, 13) = False Then
        CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    End If
    If GetFeature(hUsbDevice, array_out, 2) = False Then
        CloseHandle (hUsbDevice): ret = -94: GoTo error_exit
    End If
    CloseHandle (hUsbDevice)
    If array_out(0) <> 83 Then
           sub_YRead = -83
    End If
        OutData = array_out(1)
    Exit Function
error_exit:
    sub_YRead = ret
End Function
Public Function YWrite(ByVal indata As Byte, ByVal address As Integer, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim ary1(0 To 8) As Byte

    If (address > 495) Or (address < 0) Then YWrite = -81: Exit Function
    myconvert HKey, LKey, ary1
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    YWrite = sub_YWrite(indata, address, ary1, Path)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function
Private Function sub_YWrite(ByVal indata As Byte, ByVal address As Integer, ByRef Password() As Byte, ByVal Path As String) As Integer
    Dim addr_l As Integer
    Dim addr_h As Integer
    Dim n As Integer
    Dim ret As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
    Dim hUsbDevice As Long
    Dim opcode As Byte
    If (address > 511) Or (address < 0) Then ret = -81: GoTo error_exit
    opcode = 64
    If address > 255 Then
        opcode = 96
         address = address - 256
    End If
    
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
       array_in(1) = 17
        array_in(2) = opcode
        array_in(3) = address
        array_in(4) = indata
        For n = 0 To 7
            array_in(5 + n) = Password(n)
        Next
    If SetFeature(hUsbDevice, array_in, 13) = False Then
        CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    End If
    If GetFeature(hUsbDevice, array_out, 2) = False Then
        CloseHandle (hUsbDevice): ret = -94: GoTo error_exit
    End If
    CloseHandle (hUsbDevice)
    If array_out(1) <> 1 Then
          sub_YWrite = -82
    End If
    Exit Function
error_exit:
    sub_YWrite = ret
End Function
Public Function SetReadPassword(ByVal W_HKey As String, ByVal W_LKey As String, ByVal new_HKey As String, ByVal new_LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim ary1(0 To 8) As Byte
    Dim ary2(0 To 8) As Byte
    Dim address As Integer
    myconvert W_HKey, W_LKey, ary1
    myconvert new_HKey, new_LKey, ary2
    address = 496
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    SetReadPassword = Y_Write(ary2, address, 8, ary1, Path, 0)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal

End Function


Public Function SetWritePassword(ByVal W_HKey As String, ByVal W_LKey As String, ByVal new_HKey As String, ByVal new_LKey As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim ary1(0 To 8) As Byte
    Dim ary2(0 To 8) As Byte
    Dim address As Integer
    myconvert W_HKey, W_LKey, ary1
    myconvert new_HKey, new_LKey, ary2
    address = 504
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    SetWritePassword = Y_Write(ary2, address, 8, ary1, Path, 0)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function YWriteString(ByVal InString As String, ByVal address As Long, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
    Dim ary1(0 To 8) As Byte
    Dim hsignal As Long
    Dim n As Integer
    Dim outlen As Integer
    Dim total_len As Integer
    Dim temp_leave As Integer
    Dim leave As Integer
    Dim b() As Byte
    If (address < 0) Then YWriteString = -81: Exit Function
    myconvert HKey, LKey, ary1
    
    outlen = lstrlen(InString) '注意，这里不写入结束字符串，与原来的兼容，也可以写入结束字符串，与原来的不兼容，写入长度会增加1
    ReDim b(outlen)
    CopyStringToByte VarPtr(b(0)), InString, outlen
    
    total_len = address + outlen
    If (total_len > MAX_LEN) Then YWriteString = -47: Exit Function
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    temp_leave = address Mod 16
    leave = 16 - temp_leave
    If (leave > outlen) Then leave = outlen

    If (leave > 0) Then
        For n = 0 To (leave \ 8) - 1
            YWriteString = Y_Write(b, address + n * 8, 8, ary1, Path, n * 8)
            If (YWriteString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        Next n
        If (leave - 8 * n > 0) Then
            YWriteString = Y_Write(b, address + n * 8, (leave - n * 8), ary1, Path, 8 * n)
            If (YWriteString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        End If
    End If
    outlen = outlen - leave
    address = address + leave
    If (outlen > 0) Then
        For n = 0 To outlen \ 8 - 1
            YWriteString = Y_Write(b, address + n * 8, 8, ary1, Path, leave + n * 8)
            If (YWriteString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        Next n
        If (outlen - 8 * n > 0) Then
            YWriteString = Y_Write(b, address + n * 8, (outlen - n * 8), ary1, Path, leave + 8 * n)
            If (YWriteString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
        End If
    End If
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function YReadString(ByRef outstring As String, ByVal address As Long, ByVal nlen As Long, ByVal HKey As String, ByVal LKey As String, ByVal Path As String) As Integer
   
    Dim ary1(0 To 8) As Byte
    Dim hsignal As Long
    Dim n As Integer
    Dim total_len As Integer
    Dim outb() As Byte
    ReDim outb(nlen)
    myconvert HKey, LKey, ary1
     If (address < 0) Then YReadString = -81: Exit Function
    total_len = address + nlen
    If (total_len > MAX_LEN) Then YReadString = -47: Exit Function
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    For n = 0 To nlen \ 16 - 1
        YReadString = Y_Read(outb, address + n * 16, 16, ary1, Path, n * 16)
        If (YReadString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
    Next n
    If (nlen - 16 * n > 0) Then
        YReadString = Y_Read(outb, address + n * 16, (nlen - 16 * n), ary1, Path, 16 * n)
        If (YReadString <> 0) Then ReleaseSemaphore hsignal, 1, 0: CloseHandle hsignal: Exit Function
    End If
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
    outstring = Space(nlen)
    CopyByteToString outstring, VarPtr(outb(0)), nlen
    outstring = StringFromBuffer(outstring)
End Function

Public Function SetCal_2(ByVal Key As String, ByVal Path As String) As Integer
    Dim hsignal As Long
    Dim KeyBuf() As Byte
    Dim inb(8) As Byte
    HexStringToByteArray Key, KeyBuf
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    SetCal_2 = NT_SetCal_2(KeyBuf, 0, Path, 8)
    If (SetCal_2 <> 0) Then GoTo error1
    SetCal_2 = NT_SetCal_2(KeyBuf, 1, Path, 0)
error1:
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function Cal(ByRef InBuf() As Byte, ByRef OutBuf() As Byte, ByVal Path As String) As Integer
    Dim hsignal As Long
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    Cal = NT_Cal(InBuf, OutBuf, Path, 0)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle hsignal
End Function

Public Function EncString(ByVal InString As String, ByRef outstring As String, ByVal Path As String) As Integer
Dim hsignal As Long
Dim b() As Byte
Dim outb() As Byte
Dim n As Long
Dim nlen As Long
nlen = StringToByte(b, InString)
ReDim outb(nlen)

For n = 0 To nlen
    outb(n) = b(n)
Next n
hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
WaitForSingleObject hsignal, INFINITE
For n = 0 To nlen - 8 Step 8
    EncString = NT_Cal(b, outb, Path, n)
    If EncString <> 0 Then Exit For
Next n
ReleaseSemaphore hsignal, 1, 0
CloseHandle hsignal
outstring = ByteArrayToHexString(outb, nlen)

End Function

Public Function ReSet(ByVal Path As String) As Integer
    Dim ret As Integer
    Dim hsignal As Long
    hsignal = CreateSemaphore(0, 1, 1, "ex_sim")
    WaitForSingleObject hsignal, INFINITE
    ReSet = NT_ReSet(Path)
    ReleaseSemaphore hsignal, 1, 0
    CloseHandle (hsignal)
End Function
Private Function NT_ReSet(ByVal Path As String) As Integer
    Dim ret As Integer
    Dim array_in(0 To 25) As Byte
    Dim array_out(0 To 25) As Byte
     Dim hUsbDevice As Long
    If OpenMydivece(hUsbDevice, Path) <> 0 Then ret = -92: GoTo error_exit
    array_in(1) = 32
    If SetFeature(hUsbDevice, array_in, 2) = False Then
        CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    End If
    If GetFeature(hUsbDevice, array_out, 2) = False Then
        CloseHandle (hUsbDevice): ret = -93: GoTo error_exit
    End If
    CloseHandle (hUsbDevice)
   If array_out(0) <> 0 Then
        NT_ReSet = -82
   End If
    Exit Function
error_exit:
    NT_ReSet = ret
End Function


Private Function ReadStringEx(ByVal addr As Integer, ByRef outstring As String, ByVal DevicePath As String) As Integer
    Dim nlen As Integer
    Dim buf(1) As Byte
    '先从地址0读到以前写入的字符串的长度
    ReadStringEx = YReadEx(buf, addr, 1, "984B859C", "773F7E26", DevicePath)
    nlen = buf(0)
    If ReadStringEx <> 0 Then Exit Function
    outstring = Space(nlen)
    '再读取相应长度的字符串
    ReadStringEx = YReadString(outstring, addr + 1, nlen, "984B859C", "773F7E26", DevicePath)

End Function

Public Function CheckKeyByReadEprom() As Integer
    Dim n As Integer
    Dim DevicePath As String '用于储存加密锁所在的路径
    Dim outstring As String
    '@NoUseCode_data CheckKeyByReadEprom= 1:exit function'如果没有使用这个功能，直接返回1
    For n = 0 To 255
        CheckKeyByReadEprom = FindPort(n, DevicePath)
        If CheckKeyByReadEprom <> 0 Then Exit Function
        CheckKeyByReadEprom = ReadStringEx(0, outstring, DevicePath)
        If (CheckKeyByReadEprom = 0) And (outstring = "1031") Then Exit Function
    Next n
    CheckKeyByReadEprom = -92
End Function


Public Function CheckKeyByEncstring() As Integer
'推荐加密方案：生成随机数，让锁做加密运算，同时在程序中端使用代码做同样的加密运算，然后进行比较判断。
    
    Dim n As Integer
    Dim DevicePath As String '用于储存加密锁所在的路径
    Dim InString As String
    
    '@NoUseKeyEx CheckKeyByEncstring = 1: Exit Function '如果没有使用这个功能，直接返回1
    Randomize

    InString = Hex(CInt(Int(32767 * Rnd()))) + Hex(CInt(Int(32767 * Rnd())))

    For n = 0 To 255
        CheckKeyByEncstring = FindPort(n, DevicePath)
        If (CheckKeyByEncstring <> 0) Then Exit Function
        If (Sub_CheckKeyByEncstring(InString, DevicePath) = 0) Then
            CheckKeyByEncstring = 0
            Exit Function
        End If
    Next n
    CheckKeyByEncstring = -92
End Function


Private Function Sub_CheckKeyByEncstring(ByVal InString As String, ByVal DevicePath As String) As Integer
    ''使用增强算法对字符串进行加密
    Dim nlen As Integer
    Dim outstring As String
    Dim outstring_2 As String
    nlen = lstrlen(InString) + 1
    If (nlen < 8) Then nlen = 8
    outstring = Space(nlen * 2)
    Sub_CheckKeyByEncstring = EncString(InString, outstring, DevicePath)
    If (Sub_CheckKeyByEncstring <> 0) Then Exit Function
    outstring_2 = StrEnc(InString, "04BBE1F9D39FDA47768313D105775D1F")
    If outstring_2 = outstring Then '比较结果是否相符
        Sub_CheckKeyByEncstring = 0
    Else
        Sub_CheckKeyByEncstring = -92
    End If
End Function







