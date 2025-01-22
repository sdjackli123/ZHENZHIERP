Attribute VB_Name = "plc"
 
 Public PLCText As String  '������յ���Ϣ
 Dim ErrorPlc As String '���ؽ����Ƿ�ɹ���Ϣ
 Public SenData As String
 
Public Function DecToBin(Dats As Integer) As String                              '*ת���ɶ�����
   Const Bins = "0000000100100011010001010110011110001001101010111100110111101111"
   Dim i As Integer, s As String, Y As String
   
  Y = Hex(Dats)
  s = ""
  For i = 1 To Len(Y)
      s = s + Mid(Bins, (Val("&h" + Mid(Y, i, 1)) * 4 + 1), 4)
  Next
  DecToBin = Format(s, "00000000")                                                  '*��ʽ�����.
End Function

Private Function SumChk(Dats As String) As String
          '����У���(��λ�ַ���ASCIIֵ���)
    Dim i
    Dim CHK
    For i = 1 To Len(Dats)
             'Len �����������ַ������ַ�����Ŀ
        CHK = CHK + Asc(Mid(Dats, i, 1))
            'Asc���������ַ�ת����ASCII�롣Mid����������ߵ�iλȡһ���ַ�
    Next i
    SumChk = Right("00" + Hex(CHK), 2)
            'Hex������ת����ʮ�����ơ�Right���������ұ�ȡ��λ(Э��涨ֻȡУ��͵ĵ���λ)
End Function

Public Function OpenComm(Com As Object, ComNum As Integer, Getd As String) As String   '�򿪴���

    On Error GoTo Prog_err:
       '*�˴����ã������ѡ���˵����в����ڵ�ͨѶ�ڣ���Prog_err�������,��ʾ����Ч��ͨѶ�ڡ�
 
    If Com.PortOpen = True Then Com.PortOpen = False
             '*�����Ǵ�״̬����ر�,���д��ڵ����ù���
 
    Com.CommPort = ComNum
         '*����ͨѶ�˿ں�
    Com.Settings = Getd
       '*�趨ͨѶ��ʽ

    Com.InputLen = 0 '�����ջ���������ȫ��������
    Com.OutBufferCount = 0
       '*���ò����ط��ͻ��������ֽ���,��Ϊ0ʱ��շ��ͻ�����
    Com.InBufferCount = 0
        '*���ò����ؽ��ջ��������ֽ���,��Ϊ0ʱ��ս��ջ�����
    Com.RThreshold = 1
         '*����ON_COMMM�¼����ַ���
    Com.PortOpen = True
        '*�򿪴���
    OpenComm = "0"
    Exit Function
Prog_err:
    OpenComm = "1"
End Function

Public Function CloseComm(Com As Object) As String   '�رմ���

 On Error GoTo Prog_err:
       '*�˴����ã������ѡ���˵����в����ڵ�ͨѶ�ڣ���Prog_err�������,��ʾ����Ч��ͨѶ�ڡ�
 If Com.PortOpen = True Then
    Com.PortOpen = False
 End If
     CloseComm = "0"
    Exit Function
Prog_err:
    CloseComm = "1"
    
End Function


Public Function MSCONComm(Com As Object) As String    '������Ϣ

 Dim Getd As String                                             '*��ȡ���ջ���������
 Dim Wr(0) As String

    If Com.CommEvent = comEvReceive Then                       '*CommEvent�����Է��ص�ֵΪcomEvReceiveʱ�Ƿ����˽����¼�.
    
       Getd = Com.Input                                     '*��ȡ���ջ���������
       PLCText = PLCText & Getd
       If InStr(PLCText, Chr(21)) <> 0 Then
          '��λ������PLC����Ϣ�д��󣬲���������Ϣ���ȴ�ͨѶ��ʱ�жϣ����½���ͨѶ
          ErrorPlc = "PLC������Ϣ����"
       Else
          If InStr(PLCText, Chr(6)) <> 0 Then
             '��д�������ص���Ϣ
             ErrorPlc = "0"
             Wr(0) = "OK"
             GetData = Wr
          Else
             If InStr(PLCText, Chr(2)) = 0 Then
                 '���PLC���ص��ַ�û��Chr(2)����λ������PLC����Ϣ�д�����ǻ�û���յ���ʼ��
                ErrorPlc = "û���յ���ʼ��"
             Else
                If InStr(PLCText, Chr(3)) <> 0 And Len(PLCText) - InStr(PLCText, Chr(3)) >= 2 Then   '��Ϊ���������滹�������ַ���У���
                    '�жϽ��յ�������  ���ҽ��ճ��ȷ��Ϲ淶 ��ʼ����
                   If Mid(PLCText, InStr(PLCText, Chr(3)) + 1, 2) <> SumChk(Mid(PLCText, InStr(PLCText, Chr(2)) + 1, InStr(PLCText, Chr(3)) - InStr(PLCText, Chr(2)))) Then
                       '����У��ͼ���   ��ȥ��ʼ���������ַ���У��ͣ������Ľ���У��ͼ���
                      ErrorPlc = "У��ʹ���"
                   Else
                      PLCText = Mid(PLCText, InStr(PLCText, Chr(2)) + 1, InStr(PLCText, Chr(3)) - InStr(PLCText, Chr(2)) - 1)
                      'ȡ����Ҫ������  PLC���ص���Ϣ����ʼ��  ����  ������ У���
                      ErrorPlc = "0"

                   End If
                Else
                   ErrorPlc = "PLC��Ϣ������"
                End If
             End If
          End If
       End If
   End If
Backplc:
       MSCONComm = ErrorPlc
End Function

Public Function gk528ReadDevice(part As String, Number As Integer) As String      ' ��
      '�ַ��ӵ�ַ������
 
 Dim k As Integer
 Dim addreP As String  'λԪ��ͨѶ��ַ    �������ĸ��ַ�
 Dim DCT As String
 Dim DCTadree As String
 Dim ByteNum As Integer    '��ȡ�ֽ���
 Dim ByteStr As String  '��ȡ�ֽ��� ���Ͳ���   �����������ַ�
 
 On Error GoTo Prog_err

      '  ������ĸ������
    For k = 1 To Len(part)
        If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  ָ�����ʽ���������Ƿ�Ϊ����
           DCTadree = DCTadree & Mid(part, k, 1)  'ȡ��ַ
        Else
           DCT = DCT & Mid(part, k, 1)   'ȡ����
        End If
    Next k
    
    If Number Mod 8 > 0 Then    ' 'һ���ֽ��ǰ˸�λ����������������ȡ�ֽ������Ƕ�ȡ������8 �� 1
       ByteNum = Number \ 8 + 1
    Else
       ByteNum = Number \ 8
    End If
           
    Select Case UCase(DCT)      '  UCase  �����Сд��ĸ����ת��Ϊ��д
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
                ByteNum = Number * 2   ' һ���������ֽ�
           Case "CN"
                addreP = Right("0000" + CStr(Hex(2560 + DCTadree * 2)), 4)
                ByteNum = Number * 2   ' һ���������ֽ�
           Case "TN"
                addreP = Right("0000" + CStr(Hex(2048 + DCTadree * 2)), 4)
                ByteNum = Number * 2   ' һ���������ֽ�
   End Select
            
   ByteStr = Right("00" + CStr(Hex(ByteNum)), 2) '��ȡ�ֽڱ��뱣֤�����ַ�
   SenData = "0" + addreP + ByteStr + Chr(3) '������  λԪ����ַ   ��ȡ�ֽ���  ������
   SenData = Chr(2) + SenData + SumChk(SenData)  '��ʼ��    �����ַ���  У���

   gk528ReadDevice = "0"   '���ͳɹ������� 0
   Exit Function
Prog_err:
    gk528ReadDevice = "1"
End Function

Public Function gk528SetDevice(part As String, Number As Integer) As String    ' ��/��λ
 
 Dim k As Integer
 Dim addreP As String  'λԪ��ͨѶ��ַ    �������ĸ��ַ�
 Dim DCT As String
 Dim DCTadree As String
 Dim Func As String
   
 On Error GoTo Prog_err

      '  ������ĸ������
    For k = 1 To Len(part)
        If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  ָ�����ʽ���������Ƿ�Ϊ����
           DCTadree = DCTadree & Mid(part, k, 1)  'ȡ��ַ
        Else
           DCT = DCT & Mid(part, k, 1)   'ȡ����
        End If
    Next k

    If Number = 1 Then   '��λ������Ϊ 7  ��λ������Ϊ 8
       Func = "7"
    Else
       Func = "8"
    End If
           
    Select Case UCase(DCT)      '  UCase  �����Сд��ĸ����ת��Ϊ��д
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
           
    addreP = Right(addreP, 2) + Mid(addreP, 1, 2)  'Ԫ��ͨѶ��ַҪ�� �ߵ��ֽڻ���
    SenData = Func + addreP + Chr(3) '��/��λ����  λԪ����ַ     ������
    SenData = Chr(2) + SenData + SumChk(SenData) '��ʼ��    �����ַ���  У���
    gk528SetDevice = "0"   '������ȷ������ 0
    Exit Function
Prog_err:
     gk528SetDevice = "1"
End Function

Public Function gk528WriteDevice(part As String, Number As Integer, WriteData() As String) As String    'д

 Dim k As Integer
 Dim addreP As String  'λԪ��ͨѶ��ַ    �������ĸ��ַ�
 Dim DCT As String
 Dim DCTadree As String

 Dim WWrite() As String
 Dim j As Integer
 Dim WriteD As String  'д����ֵ
 Dim Write2 As String '�ۺ�д������
 
 On Error GoTo Prog_err

      '  ������ĸ������
     For k = 1 To Len(part)
         If IsNumeric(Mid(part, k, 1)) = True Then     'IsNumeric  ָ�����ʽ���������Ƿ�Ϊ����
            DCTadree = DCTadree & Mid(part, k, 1)  'ȡ��ַ
         Else
            DCT = DCT & Mid(part, k, 1)   'ȡ����
         End If
     Next k
       
     ByteNum = Number * 2   ' һ���������ֽ�
     ByteStr = Right("00" + CStr(Hex(ByteNum)), 2)  'д���ֽڱ��뱣֤�����ַ�
     
     Select Case UCase(DCT)      '  UCase  �����Сд��ĸ����ת��Ϊ��д
            Case "D"
                 addreP = Right("0000" + CStr(Hex(4096 + Val(DCTadree) * 2)), 4)
            Case "CN"
                 addreP = Right("0000" + CStr(Hex(2560 + Val(DCTadree) * 2)), 4)
            Case "TN"
                 addreP = Right("0000" + CStr(Hex(2048 + Val(DCTadree) * 2)), 4)
     End Select
      ReDim WWrite(0 To Number - 1) As String   '����д��������¶��� д����ֵ����
      For j = 0 To Number - 1
          WriteD = Right("0000" + Hex(Val(WriteData(j))), 4)
          WWrite(j) = Right(WriteD, 2) + Mid(WriteD, 1, 2)   'д�����ݱ���  ʮ������  ÿ���ֱ�֤�ĸ��ַ�
          Write2 = Write2 + WWrite(j)
      Next j
      SenData = "1" + addreP + ByteStr + Write2 + Chr(3) 'д����  λԪ����ַ   д���ֽ���  д����ֵ  ������
      SenData = Chr(2) + SenData + SumChk(SenData) '��ʼ��    �����ַ���  У���
      gk528WriteDevice = "0"   '���ͳɹ������� 0
     Exit Function
Prog_err:
     gk528WriteDevice = "1"
End Function



