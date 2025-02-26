Attribute VB_Name = "Module2"
Private Const MAX_IDE_DRIVES As Long = 4   ' Max number of drives assuming primary/secondary, master/slave topology
Private Const READ_ATTRIBUTE_BUFFER_SIZE As Long = 512
Private Const IDENTIFY_BUFFER_SIZE As Long = 512
Private Const READ_THRESHOLD_BUFFER_SIZE As Long = 512
Private Const DFP_GET_VERSION As Long = &H74080
Private Const DFP_SEND_DRIVE_COMMAND As Long = &H7C084
Private Const DFP_RECEIVE_DRIVE_DATA As Long = &H7C088

Private Type GETVERSIONOUTPARAMS
    bVersion As Byte       ' Binary driver version.
    bRevision As Byte      ' Binary driver revision.
    bReserved As Byte      ' Not used.
    bIDEDeviceMap As Byte  ' Bit map of IDE devices.
    fCapabilities As Long  ' Bit mask of driver capabilities.
    dwReserved(3) As Long  ' For future use.
End Type

Private Const CAP_IDE_ID_FUNCTION As Long = 1               ' ATA ID command supported
Private Const CAP_IDE_ATAPI_ID As Long = 2                  ' ATAPI ID command supported
Private Const CAP_IDE_EXECUTE_SMART_FUNCTION As Long = 4    ' SMART commannds supported

Private Type IDEREGS
    bFeaturesReg As Byte       ' Used for specifying SMART "commands".
    bSectorCountReg As Byte    ' IDE sector count register
    bSectorNumberReg As Byte   ' IDE sector number register
    bCylLowReg As Byte         ' IDE low order cylinder value
    bCylHighReg As Byte        ' IDE high order cylinder value
    bDriveHeadReg As Byte      ' IDE drive/head register
    bCommandReg As Byte        ' Actual IDE command.
    bReserved As Byte          ' reserved for future use.  Must be zero.
End Type

Private Type SENDCMDINPARAMS
    cBufferSize As Long        ' Buffer size in bytes
    irDriveRegs As IDEREGS     ' Structure with drive register values.
    bDriveNumber As Byte       ' Physical drive number to send
    ' command to (0,1,2,3).
    bReserved(2) As Byte       ' Reserved for future expansion.
    dwReserved(3) As Long      ' For future use.
    bBuffer(0) As Byte         ' Input buffer.
End Type

Private Const IDE_ATAPI_ID As Long = &HA1  ' Returns ID sector for ATAPI.
Private Const IDE_ID_FUNCTION As Long = &HEC  ' Returns ID sector for ATA.
Private Const IDE_EXECUTE_SMART_FUNCTION As Long = &HB0  ' Performs SMART cmd.
Private Const SMART_CYL_LOW As Long = &H4F
Private Const SMART_CYL_HI As Long = &HC2

Private Type DRIVERSTATUS
    bDriverError As Byte       ' Error code from driver,
    bIDEStatus As Byte         ' Contents of IDE Error register.
    bReserved(1) As Byte       ' Reserved for future expansion.
    dwReserved(1) As Long      ' Reserved for future expansion.
End Type

Private Const SMART_NO_ERROR As Long = 0  ' No error
Private Const SMART_IDE_ERROR As Long = 1  ' Error from IDE controller
Private Const SMART_INVALID_FLAG As Long = 2  ' Invalid command flag
Private Const SMART_INVALID_COMMAND As Long = 3  ' Invalid command byte
Private Const SMART_INVALID_BUFFER As Long = 4  ' Bad buffer (null, invalid addr..)
Private Const SMART_INVALID_DRIVE As Long = 5  ' Drive number not valid
Private Const SMART_INVALID_IOCTL As Long = 6   ' Invalid IOCTL
Private Const SMART_ERROR_NO_MEM As Long = 7  ' Could not lock user's buffer
Private Const SMART_INVALID_REGISTER As Long = 8  ' Some IDE Register not valid
Private Const SMART_NOT_SUPPORTED As Long = 9  ' Invalid cmd flag set
Private Const SMART_NO_IDE_DEVICE As Long = 10 ' Cmd issued to device not present

Private Type SENDCMDOUTPARAMS
    cBufferSize As Long        ' Size of bBuffer in bytes
    drvStatus As DRIVERSTATUS  ' Driver status structure.
    bBuffer(0) As Byte         ' Buffer of arbitrary length in which to store the data read from the                                          ' drive.
End Type

Private Const SMART_READ_ATTRIBUTE_VALUES As Long = &HD0    ' ATA4: Renamed
Private Const SMART_READ_ATTRIBUTE_THRESHOLDS As Long = &HD1    ' Obsoleted in ATA4!
Private Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE As Long = &HD2
Private Const SMART_SAVE_ATTRIBUTE_VALUES As Long = &HD3
Private Const SMART_EXECUTE_OFFLINE_IMMEDIATE As Long = &HD4    ' ATA4
Private Const SMART_ENABLE_SMART_OPERATIONS As Long = &HD8
Private Const SMART_DISABLE_SMART_OPERATIONS As Long = &HD9
Private Const SMART_RETURN_SMART_STATUS As Long = &HDA

Private Type DRIVEATTRIBUTE
    bAttrID As Byte        ' Identifies which attribute
    wStatusFlags As Integer    ' see bit definitions below
    bAttrValue As Byte     ' Current normalized value
    bWorstValue As Byte    ' How bad has it ever been?
    bRawValue(5) As Byte   ' Un-normalized value
    bReserved As Byte      ' ...
End Type

Private Type ATTRTHRESHOLD
    bAttrID As Byte            ' Identifies which attribute
    bWarrantyThreshold As Byte ' Triggering value
    bReserved(9) As Byte      ' ...
End Type

Private Type IDSECTOR
    wGenConfig As Integer
    wNumCyls As Integer
    wReserved As Integer
    wNumHeads As Integer
    wBytesPerTrack As Integer
    wBytesPerSector As Integer
    wSectorsPerTrack As Integer
    wVendorUnique(2) As Integer
    sSerialNumber(19) As Byte
    wBufferType As Integer
    wBufferSize As Integer
    wECCSize As Integer
    sFirmwareRev(7) As Byte
    sModelNumber(39) As Byte
    wMoreVendorUnique As Integer
    wDoubleWordIO As Integer
    wCapabilities As Integer
    wReserved1 As Integer
    wPIOTiming As Integer
    wDMATiming As Integer
    wBS As Integer
    wNumCurrentCyls As Integer
    wNumCurrentHeads As Integer
    wNumCurrentSectorsPerTrack As Integer
    ulCurrentSectorCapacity(3) As Byte    '这里只能用byte，因为VB没有无符号的LONG型变量
    wMultSectorStuff As Integer
    ulTotalAddressableSectors(3) As Byte   '这里只能用byte，因为VB没有无符号的LONG型变量
    wSingleWordDMA As Integer
    wMultiWordDMA As Integer
    bReserved(127) As Byte
End Type

Private Const ATTR_INVALID As Long = 0
Private Const ATTR_READ_ERROR_RATE As Long = 1
Private Const ATTR_THROUGHPUT_PERF As Long = 2
Private Const ATTR_SPIN_UP_TIME As Long = 3
Private Const ATTR_START_STOP_COUNT As Long = 4
Private Const ATTR_REALLOC_SECTOR_COUNT As Long = 5
Private Const ATTR_READ_CHANNEL_MARGIN As Long = 6
Private Const ATTR_SEEK_ERROR_RATE As Long = 7
Private Const ATTR_SEEK_TIME_PERF As Long = 8
Private Const ATTR_POWER_ON_HRS_COUNT As Long = 9
Private Const ATTR_SPIN_RETRY_COUNT As Long = 10
Private Const ATTR_CALIBRATION_RETRY_COUNT As Long = 11
Private Const ATTR_POWER_CYCLE_COUNT As Long = 12

Private Const PRE_FAILURE_WARRANTY As Long = &H1
Private Const ON_LINE_COLLECTION As Long = &H2
Private Const PERFORMANCE_ATTRIBUTE As Long = &H4
Private Const ERROR_RATE_ATTRIBUTE As Long = &H8
Private Const EVENT_COUNT_ATTRIBUTE As Long = &H10
Private Const SELF_PRESERVING_ATTRIBUTE As Long = &H20

Private Const NUM_ATTRIBUTE_STRUCTS As Long = 30
Private Const INVALID_HANDLE_VALUE As Long = -1

Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Const CREATE_NEW As Long = 1
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING  As Long = 3

Private m_DiskInfo As IDSECTOR

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'信息类型枚举
Enum eumInfoType
    hdmodelsn = 0
    hdOnlyModel = 1
    hdOnlySN = 2
End Enum

'磁盘通道枚举
Enum eumDiskNo
    hdPrimaryMaster = 0
    hdPrimarySlave = 1
    hdSecondaryMaster = 2
    hdSecondarySlave = 3
End Enum

'取得逻辑盘序列号（非唯一）
Function GetDiskVolume(Optional ByVal strDiskName = "C") As String
    Dim TempStr1 As String * 256, TempStr2 As String * 256
    Dim TempLon1 As Long, TempLon2 As Long, GetVal As Long
    
    Dim tmpVol As String
    
    Call GetVolumeInformation(strDiskName & ":\", TempStr1, 256, GetVal, TempLon1, TempLon2, TempStr2, 256)
    If GetVal = 0 Then
        tmpVol = ""
    Else
        tmpVol = Hex(GetVal)
        tmpVol = String(8 - Len(tmpVol), "0") & tmpVol
        tmpVol = Left(tmpVol, 4) & "-" & Right(tmpVol, 4)
    End If
    GetDiskVolume = tmpVol
    DISKCO = tmpVol
End Function

'取得硬盘信息：型号/物理系列号（唯一）
Function GetHardDiskInfo(Optional ByVal numDisk As eumDiskNo = hdPrimaryMaster, Optional ByVal numType As eumInfoType = hdOnlySN) As String

    If GetDiskInfo(numDisk) = 1 Then
        Dim pSerialNumber As String, pModelNumber As String
        pSerialNumber = StrConv(m_DiskInfo.sSerialNumber, vbUnicode)
        pModelNumber = StrConv(m_DiskInfo.sModelNumber, vbUnicode)
        
        Select Case numType
            Case hdOnlyModel  '仅型号
                GetHardDiskInfo = Trim(pModelNumber)
                DISKNO = Trim(pModelNumber)
            Case hdOnlySN  '仅系列号
                GetHardDiskInfo = Trim(pSerialNumber)
                DISKNO = Trim(pSerialNumber)
            Case Else   '型号,系列号
                GetHardDiskInfo = Trim(pModelNumber) & "," & Trim(pSerialNumber)
               DISKNO = Trim(pModelNumber) & "+" & Trim(pSerialNumber)
        End Select
     End If

End Function

Private Function OpenSMART(ByVal nDrive As Byte) As Long
  Dim hSMARTIOCTL As Long
  Dim hd As String
  Dim VersionInfo As OSVERSIONINFO

    hSMARTIOCTL = INVALID_HANDLE_VALUE
    VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
    GetVersionEx VersionInfo
    Select Case VersionInfo.dwPlatformId
      Case VER_PLATFORM_WIN32s
        OpenSMART = hSMARTIOCTL
      Case VER_PLATFORM_WIN32_WINDOWS
        hSMARTIOCTL = CreateFile("\\.\SMARTVSD", 0, 0, 0, CREATE_NEW, 0, 0)
      Case VER_PLATFORM_WIN32_NT
        If nDrive < MAX_IDE_DRIVES Then
            hd = "\\.\PhysicalDrive" & nDrive
            hSMARTIOCTL = CreateFile(hd, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
        End If
    End Select
    OpenSMART = hSMARTIOCTL

End Function

Private Function DoIDENTIFY(ByVal hSMARTIOCTL As Long, pSCIP As SENDCMDINPARAMS, pSCOP() As Byte, ByVal bIDCmd As Byte, ByVal bDriveNum As Byte, lpcbBytesReturned As Long) As Boolean
    pSCIP.cBufferSize = IDENTIFY_BUFFER_SIZE

    pSCIP.irDriveRegs.bFeaturesReg = 0
    pSCIP.irDriveRegs.bSectorCountReg = 1
    pSCIP.irDriveRegs.bSectorNumberReg = 1
    pSCIP.irDriveRegs.bCylLowReg = 0
    pSCIP.irDriveRegs.bCylHighReg = 0

    pSCIP.irDriveRegs.bDriveHeadReg = &HA0 Or ((bDriveNum And 1) * 2 ^ 4)
    '
    pSCIP.irDriveRegs.bCommandReg = bIDCmd
    pSCIP.bDriveNumber = bDriveNum
    pSCIP.cBufferSize = IDENTIFY_BUFFER_SIZE
   DoIDENTIFY = CBool(DeviceIoControl(hSMARTIOCTL, DFP_RECEIVE_DRIVE_DATA, _
                 pSCIP, 32, _
                 pSCOP(0), 528, _
                 lpcbBytesReturned, 0))

End Function

Private Function DoEnableSMART(ByVal hSMARTIOCTL As Long, pSCIP As SENDCMDINPARAMS, pSCOP As SENDCMDOUTPARAMS, ByVal bDriveNum As Byte, lpcbBytesReturned As Long) As Boolean
    pSCIP.cBufferSize = 0

    pSCIP.irDriveRegs.bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
    pSCIP.irDriveRegs.bSectorCountReg = 1
    pSCIP.irDriveRegs.bSectorNumberReg = 1
    pSCIP.irDriveRegs.bCylLowReg = SMART_CYL_LOW
    pSCIP.irDriveRegs.bCylHighReg = SMART_CYL_HI
    pSCIP.irDriveRegs.bDriveHeadReg = &HA0 Or ((bDriveNum And 1) * 2 ^ 4)
    pSCIP.irDriveRegs.bCommandReg = IDE_EXECUTE_SMART_FUNCTION
    pSCIP.bDriveNumber = bDriveNum

    DoEnableSMART = CBool(DeviceIoControl(hSMARTIOCTL, DFP_SEND_DRIVE_COMMAND, _
                    pSCIP, LenB(pSCIP) - 1, _
                    pSCOP, LenB(pSCOP) - 1, _
                    lpcbBytesReturned, 0))

End Function

'---------------------------------------------------------------------
'---------------------------------------------------------------------
Private Sub ChangeByteOrder(szString() As Byte, ByVal uscStrSize As Integer)

  Dim i As Integer
  Dim bTemp As Byte

    For i = 0 To uscStrSize - 1 Step 2
        bTemp = szString(i)
        szString(i) = szString(i + 1)
        szString(i + 1) = bTemp
    Next i

End Sub

Private Sub DisplayIdInfo(pids As IDSECTOR, pSCIP As SENDCMDINPARAMS, ByVal bIDCmd As Byte, ByVal bDfpDriveMap As Byte, ByVal bDriveNum As Byte)

    ChangeByteOrder pids.sModelNumber, UBound(pids.sModelNumber) + 1

    ChangeByteOrder pids.sFirmwareRev, UBound(pids.sFirmwareRev) + 1

    ChangeByteOrder pids.sSerialNumber, UBound(pids.sSerialNumber) + 1

End Sub

Public Function GetDiskInfo(ByVal nDrive As Byte) As Long

  Dim hSMARTIOCTL As Long
  Dim cbBytesReturned As Long
  Dim VersionParams As GETVERSIONOUTPARAMS
  Dim scip As SENDCMDINPARAMS
  Dim scop() As Byte
  Dim OutCmd As SENDCMDOUTPARAMS
  Dim bDfpDriveMap As Byte
  Dim bIDCmd As Byte                    ' IDE or ATAPI IDENTIFY cmd
  Dim uDisk As IDSECTOR

    m_DiskInfo = uDisk
    '
    '
    hSMARTIOCTL = OpenSMART(nDrive)
    If hSMARTIOCTL <> INVALID_HANDLE_VALUE Then

        Call DeviceIoControl(hSMARTIOCTL, DFP_GET_VERSION, ByVal 0, 0, VersionParams, Len(VersionParams), cbBytesReturned, 0)

        If Not (VersionParams.bIDEDeviceMap \ 2 ^ nDrive And &H10) Then
            If DoEnableSMART(hSMARTIOCTL, scip, OutCmd, nDrive, cbBytesReturned) Then
                bDfpDriveMap = bDfpDriveMap Or 2 ^ nDrive
            End If
        End If
        bIDCmd = IIf((VersionParams.bIDEDeviceMap \ 2 ^ nDrive And &H10), IDE_ATAPI_ID, IDE_ID_FUNCTION)

        ReDim scop(LenB(OutCmd) + IDENTIFY_BUFFER_SIZE - 1) As Byte
        If DoIDENTIFY(hSMARTIOCTL, scip, scop, bIDCmd, nDrive, cbBytesReturned) Then
            CopyMemory m_DiskInfo, scop(LenB(OutCmd) - 4), LenB(m_DiskInfo)
            Call DisplayIdInfo(m_DiskInfo, scip, bIDCmd, bDfpDriveMap, nDrive)
            CloseHandle hSMARTIOCTL
            GetDiskInfo = 1
            Exit Function '>---> Bottom
        End If
        CloseHandle hSMARTIOCTL
        GetDiskInfo = 0
      Else 'NOT HSMARTIOCTL...
        GetDiskInfo = -1
    End If

End Function




