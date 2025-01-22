Attribute VB_Name = "plc"
 
 Public PLCText As String  '保存接收的信息
 Dim ErrorPlc As String '返回接收是否成功信息
 Public SenData As String
 
Public Function DecToBin(Dats As Integer) As String                              '*转换成二进制
   Const Bins = "0000000100100011010001010110011110001001101010111100110111101111"
   Dim i As Integer, s As String, Y As String
   
  Y = Hex(Dats)
  s = ""
  For i = 1 To Len(Y)
      s = s + Mid(Bins, (Val("&h" + Mid(Y, i, 1)) * 4 + 1), 4)
  Next
  DecToBin = Format(s, "00000000")                                                  '*格式化输出.
End Function

Private Function SumChk(Dats As String) As String
          '计算校验和(各位字符的ASCII值相加)
    Dim i
    Dim CHK
    For i = 1 To Len(Dats)
             'Len 函数，返回字符串内字符的数目
        CHK = CHK + Asc(Mid(Dats, i, 1))
            'Asc函数，将字符转换成ASCII码。Mid函数，从左边第i位取一个字符
    Next i
    SumChk = Right("00" + Hex(CHK), 2)
            'Hex函数，转换成十六进制。Right函数，从右边取两位(协议规定只取校验和的低两位)
End Function

Public Function OpenComm(Com As Object, ComNum As Integer, Getd As String) As String   '打开串口

    On Error GoTo Prog_err:
       '*此处作用：如果您选择了电脑中不存在的通讯口，则‘Prog_err’程序段,提示“无效的通讯口”
 
    If Com.PortOpen = True Then Com.PortOpen = False
             '*串口是打开状态，则关闭,进行串口的设置工作
 
    Com.CommPort = ComNum
         '*设置通讯端口号
    Com.Settings = Getd
       '*设定通讯格式

    Com.InputLen = 0 '将接收缓冲区内容全部读回来
    Com.OutBufferCount = 0
       '*设置并返回发送缓冲区的字节数,设为0时清空发送缓冲区
    Com.InBufferCount = 0
        '*设置并返回接收缓冲区的字节数,设为0时清空接收缓冲区
    Com.RThreshold = 1
         '*产生ON_COMMM事件的字符数
    Com.PortOpen = True
        '*打开串口
    OpenComm = "0"
    Exit Function
Prog_err:
    OpenComm = "1"
End Function

Public Function CloseComm(Com As Object) As String   '关闭串口

 On Error GoTo Prog_err:
       '*此处作用：如果您选择了电脑中不存在的通讯口，则‘Prog_err’程序段,提示“无效的通讯口”
 If Com.PortOpen = True Then
    Com.PortOpen = False
 End If
     CloseComm = "0"
    Exit Function
Prog_err:
    CloseComm = "1"
    
End Function


Public Function MSCONComm(Com As Object) As String    '返回信息

 Dim Getd As String                                             '*读取接收缓冲区变量
 Dim Wr(0) As String

    If Com.CommEvent = comEvReceive Then                       '*CommEvent的属性返回的值为comEvReceive时是发生了接收事件.
    
       Getd = Com.Input                                     '*读取接收缓冲区内容
       PLCText = PLCText & Getd
       If InStr(PLCText, Chr(21)) <> 0 Then
          '上位机传给PLC的信息有错误，不处理返回信息，等待通讯超时判断，重新建立通讯
          ErrorPlc = "PLC发送信息错误"
       Else
          If InStr(PLCText, Chr(6)) <> 0 Then
             '是写操作返回的信息
             ErrorPlc = "0"
             Wr(0) = "OK"
             GetData = Wr
          Else
             If InStr(PLCText, Chr(2)) = 0 Then
                 '如果PLC返回的字符没有Chr(2)，上位机传给PLC的信息有错误或是还没接收到起始符
                ErrorPlc = "没接收到起始符"
             Else
                If InStr(PLCText, Chr(3)) <> 0 And Len(PLCText) - InStr(PLCText, Chr(3)) >= 2 Then   '因为结束符后面还有两个字符的校验和
                    '判断接收到结束符  并且接收长度符合规范 则开始分析
                   If Mid(PLCText, InStr(PLCText, Chr(3)) + 1, 2) <> SumChk(Mid(PLCText, InStr(PLCText, Chr(2)) + 1, InStr(PLCText, Chr(3)) - InStr(PLCText, Chr(2)))) Then
                       '接收校验和计算   除去起始符及两个字符的校验和，其他的进行校验和计算
                      ErrorPlc = "校验和错误"
                   Else
                      PLCText = Mid(PLCText, InStr(PLCText, Chr(2)) + 1, InStr(PLCText, Chr(3)) - InStr(PLCText, Chr(2)) - 1)
                      '取出需要的数据  PLC返回的信息：起始符  数据  结束符 校验和
                      ErrorPlc = "0"

                   End If
                Else
                   ErrorPlc = "PLC信息返回中"
                End If
             End If
          End If
       End If
   End If
Backplc:
       MSCONComm = ErrorPlc
End Function

Public Function gk528ReadDevice(part As String, Number As Integer) As String      ' 读
      '字符加地址，数量
 
 Dim k As Integer
 Dim addreP As String  '位元件通讯地址    必须是四个字符
 Dim DCT As String
 Dim DCTadree As String
 Dim ByteNum As Integer    '读取字节数
 Dim ByteStr As String  '读取字节数 发送部分   必须是两个字符
 
 On Error GoTo Prog_err

      '  分离字母与数字
    For k = 1 To Len(part)
        If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  指出表达式的运算结果是否为数。
           DCTadree = DCTadree & Mid(part, k, 1)  '取地址
        Else
           DCT = DCT & Mid(part, k, 1)   '取符号
        End If
    Next k
    
    If Number Mod 8 > 0 Then    ' '一个字节是八个位，如果有余数，则读取字节数，是读取个数除8 加 1
       ByteNum = Number \ 8 + 1
    Else
       ByteNum = Number \ 8
    End If
           
    Select Case UCase(DCT)      '  UCase  如果是小写字母，则转换为大写
           Case "X"
                addreP = Right("0000" + Hex(Val("&o" + DCTadree) \ 8 + 128), 4)
           Case "Y"
                addreP = Right("0000" + Hex(Val("&o" + DCTadree) \ 8 + 160), 4)
           Case "M"
                addreP = Right("0000" + Hex(Val(DCTadree) \ 8 + 256), 4)
           Case "S"
                addreP = Right("0000" + Hex(Val(DCTadree) \ 8), 4)
           Case "C"
                addreP = Right("0000" + Hex(Val(DCTadree) \ 8 + 448), 4)
           Case "T"
                addreP = Right("0000" + Hex(Val(DCTadree) \ 8 + 192), 4)
           Case "D"
                addreP = Right("0000" + CStr(Hex(4096 + DCTadree * 2)), 4)
                ByteNum = Number * 2   ' 一个字两个字节
           Case "CN"
                addreP = Right("0000" + CStr(Hex(2560 + DCTadree * 2)), 4)
                ByteNum = Number * 2   ' 一个字两个字节
           Case "TN"
                addreP = Right("0000" + CStr(Hex(2048 + DCTadree * 2)), 4)
                ByteNum = Number * 2   ' 一个字两个字节
   End Select
            
   ByteStr = Right("00" + CStr(Hex(ByteNum)), 2) '读取字节必须保证两个字符
   SenData = "0" + addreP + ByteStr + Chr(3) '读命令  位元件地址   读取字节数  结束符
   SenData = Chr(2) + SenData + SumChk(SenData)  '起始符    命令字符串  校验和

   gk528ReadDevice = "0"   '发送成功，返回 0
   Exit Function
Prog_err:
    gk528ReadDevice = "1"
End Function

Public Function gk528SetDevice(part As String, Number As Integer) As String    ' 置/复位
 
 Dim k As Integer
 Dim addreP As String  '位元件通讯地址    必须是四个字符
 Dim DCT As String
 Dim DCTadree As String
 Dim Func As String
   
 On Error GoTo Prog_err

      '  分离字母与数字
    For k = 1 To Len(part)
        If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  指出表达式的运算结果是否为数。
           DCTadree = DCTadree & Mid(part, k, 1)  '取地址
        Else
           DCT = DCT & Mid(part, k, 1)   '取符号
        End If
    Next k

    If Number = 1 Then   '置位命令码为 7  复位命令码为 8
       Func = "7"
    Else
       Func = "8"
    End If
           
    Select Case UCase(DCT)      '  UCase  如果是小写字母，则转换为大写
           Case "X"
                addreP = Right("0000" + Hex(Val("&o" + DCTadree) + 1024), 4)
           Case "Y"
                addreP = Right("0000" + Hex(Val("&o" + DCTadree) + 1280), 4)
           Case "M"
                addreP = Right("0000" + Hex(Val(DCTadree) + 2048), 4)
           Case "S"
                addreP = Right("0000" + Hex(Val(DCTadree)), 4)
           Case "C"
                addreP = Right("0000" + Hex(Val(DCTadree) + 3584), 4)
           Case "T"
                addreP = Right("0000" + Hex(Val(DCTadree) + 1536), 4)
    End Select
           
    addreP = Right(addreP, 2) + Mid(addreP, 1, 2)  '元件通讯地址要求 高低字节互换
    SenData = Func + addreP + Chr(3) '置/复位命令  位元件地址     结束符
    SenData = Chr(2) + SenData + SumChk(SenData) '起始符    命令字符串  校验和
    gk528SetDevice = "0"   '发送正确，返回 0
    Exit Function
Prog_err:
     gk528SetDevice = "1"
End Function

Public Function gk528WriteDevice(part As String, Number As Integer, WriteData() As String) As String    '写

 Dim k As Integer
 Dim addreP As String  '位元件通讯地址    必须是四个字符
 Dim DCT As String
 Dim DCTadree As String

 Dim WWrite() As String
 Dim j As Integer
 Dim WriteD As String  '写入数值
 Dim Write2 As String '综合写入数据
 
 On Error GoTo Prog_err

      '  分离字母与数字
     For k = 1 To Len(part)
         If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  指出表达式的运算结果是否为数。
            DCTadree = DCTadree & Mid(part, k, 1)  '取地址
         Else
            DCT = DCT & Mid(part, k, 1)   '取符号
         End If
     Next k
       
     ByteNum = Number * 2   ' 一个字两个字节
     ByteStr = Right("00" + CStr(Hex(ByteNum)), 2)  '写入字节必须保证两个字符
     
     Select Case UCase(DCT)      '  UCase  如果是小写字母，则转换为大写
            Case "D"
                 addreP = Right("0000" + CStr(Hex(4096 + Val(DCTadree) * 2)), 4)
            Case "CN"
                 addreP = Right("0000" + CStr(Hex(2560 + Val(DCTadree) * 2)), 4)
            Case "TN"
                 addreP = Right("0000" + CStr(Hex(2048 + Val(DCTadree) * 2)), 4)
     End Select
      ReDim WWrite(0 To Number - 1) As String   '根据写入个数重新定义 写入数值变量
      For j = 0 To Number - 1
          WriteD = Right("0000" + Hex(Val(WriteData(j))), 4)
          WWrite(j) = Right(WriteD, 2) + Mid(WriteD, 1, 2)   '写入数据必须  十六进制  每个字保证四个字符
          Write2 = Write2 + WWrite(j)
      Next j
      SenData = "1" + addreP + ByteStr + Write2 + Chr(3) '写命令  位元件地址   写入字节数  写入数值  结束符
      SenData = Chr(2) + SenData + SumChk(SenData) '起始符    命令字符串  校验和
      gk528WriteDevice = "0"   '发送成功，返回 0
     Exit Function
Prog_err:
     gk528WriteDevice = "1"
End Function



