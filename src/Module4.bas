Attribute VB_Name = "Module4"
Public Sub CPFH(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data9.RecordSource = "SELECT * FROM CLYSHZ order BY �ӹ���λ"
Data9.Refresh
Data8.RecordSource = "SELECT * FROM CLZZPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLZZPZ WHERE ���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data7.Refresh
PZH = "5-1"
If Data7.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data7.Recordset.Fields(0) + 1)
End If

If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "��Ӫ����"
Data8.Recordset.Fields(1) = "Ӧ���˿�"
Data8.Recordset.Fields(2) = Data9.Recordset.Fields(1)
Data8.Recordset.Fields(3) = "��Ӫҵ������"
Data8.Recordset.Fields(4) = ""
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(4)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = dt3
Data8.Recordset.Fields(8) = ""
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = "�Զ�"
Data8.Recordset.Update

Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("��Ʒ������ת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.ƾ֤��,3))) FROM CLZZPZ WHERE CLZZPZ.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "5-1"
If Data7.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("��Ʒ������ת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")

End If
End Sub

Public Sub FKHZXJ()   '''''''''�������--�ֽ�
On Error Resume Next

Data8.RecordSource = "SELECT * FROM CLFKPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'2-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "2-1"
If Not Data7.Recordset.EOF Then
PZH = "2-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM JGZCX1 where val(�����ָ���)>0"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "������"     '''''ժҪ
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)      ''''�跽���˿�Ŀ
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)         ''''�跽��ϸ��Ŀ
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(4)
End If
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)     '''''''''''�������˿�Ŀ
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)       '''������ϸ��Ŀ
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(0)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(6)     ''''''''�������
Data8.Recordset.Fields(6) = PZH                           '''''''''''''ƾ֤��
Data8.Recordset.Fields(7) = DTPicker3.Value               '''''''��������
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""                           ''''''''''����
Data8.Recordset.Fields(10) = ""                          ''''''''''''����
Data8.Recordset.Fields(11) = DBCombo3.Text       ''''''''''''''�Ƶ�
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "��"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ𸶿�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�ֽ𸶿"
Data2.Recordset.Fields(3) = "����ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'2-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "2-1"
If Not Data7.Recordset.EOF Then
PZH = "2-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ𸶿�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�ֽ𸶿"
Data2.Recordset.Fields(3) = "����ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
End If
End Sub
''''''''''''''
Public Sub FKHZYH()   '''''''''�������---���д��
On Error Resume Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'4-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "4-1"
If Not Data7.Recordset.EOF Then
PZH = "4-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �������<>'0' AND �跽���='0' AND INSTR(���,'���д��')>0 AND (���<>'��' OR ���=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(4)
End If
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(0)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(6)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "��"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "���и��"
Data2.Recordset.Fields(3) = "����ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'4-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "4-1"
If Not Data7.Recordset.EOF Then
PZH = "4-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "���и��"
Data2.Recordset.Fields(3) = "����ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���ƾ֤")
End If
End Sub

Public Sub SKHZXJ()    ''''''''�տ����----�ֽ�
On Error Resume Next
If DBCombo3.Text = "" Then
MsgBox ("��ѡ�񸴺�Ա")
Exit Sub
End If
Data8.RecordSource = "SELECT * FROM CLSKPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'1-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "1-1"
If Not Data7.Recordset.EOF Then
PZH = "1-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �跽���<>'0' AND �������='0' AND INSTR(���,'�ֽ�')>0 AND (���<>'��' OR ���=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(0)
End If
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(4)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(5)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "��"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ��տ�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�ֽ��տ"
Data2.Recordset.Fields(3) = "�տ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'1-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "1-1"
If Not Data7.Recordset.EOF Then
PZH = "1-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�ֽ��տ"
Data2.Recordset.Fields(3) = "�տ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ��տ�ƾ֤")
End If
End Sub

Public Sub SKHZYH()    ''''''''�տ����----���д��
On Error Resume Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'3-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "3-1"
If Not Data7.Recordset.EOF Then
PZH = "3-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �跽���<>'0' AND �������='0' AND INSTR(���,'���д��')>0 AND (���<>'��' OR ���=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(0)
End If
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(4)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(5)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "��"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���տ�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�����տ"
Data2.Recordset.Fields(3) = "�տ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'3-')>0 AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "3-1"
If Not Data7.Recordset.EOF Then
PZH = "3-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "�����տ"
Data2.Recordset.Fields(3) = "�տ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���տ�ƾ֤")
End If
End Sub



Public Sub CLRK(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data9.RecordSource = "SELECT * FROM JGZCX1 where val(����Ӧ����)>0"
Data9.Refresh
If Data9.Recordset.EOF Then Exit Sub
Data10.RecordSource = "SELECT * FROM CLZZPZ"
Data10.Refresh
Data11.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.ƾ֤��,3))) FROM CLZZPZ WHERE CLZZPZ.���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "5-1"
If Data11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data10.Recordset.AddNew
Data10.Recordset.Fields(0) = "������"
Data10.Recordset.Fields(1) = "�������"
Data10.Recordset.Fields(2) = ""
Data10.Recordset.Fields(3) = "Ӧ���˿�"
Data10.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data10.Recordset.Fields(5) = Data9.Recordset.Fields(2)
Data10.Recordset.Fields(6) = PZH
Data10.Recordset.Fields(7) = CDate(dt3)
Data10.Recordset.Fields(8) = ""
Data10.Recordset.Fields(9) = ""
Data10.Recordset.Fields(10) = ""
Data10.Recordset.Fields(11) = "�Զ�-����"
Data10.Recordset.Update
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data11.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.ƾ֤��,3))) FROM CLZZPZ WHERE CLZZPZ.���� BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "5-1"
If Data11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Loop
MsgBox ("������ⵥת�˳ɹ���" + "����" + Str(KLLLL) + "ƾ֤")
End Sub

Public Sub CLCK()
On Error Resume Next
Data8.Database.Execute "DELETE * FROM CLSCHZ"
lo = "d:\���ݿ�\ssdt\" + LJB + "\zcw.MDB"       '''''''''''''''''''''''����
Data3.Database.Execute "INSERT INTO CLSCHZ(���) IN'" & lo & "' SELECT FORMAT(SUM(�ϼƽ��),'#0.00') AS ��� FROM KPD WHERE (���<>'��' OR ���=NULL) AND ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data8.Database.Execute "DELETE * FROM CLSCHZ WHERE ���=NULL"
Data9.RecordSource = "SELECT * FROM CLSCHZ"
Data9.Refresh
Data8.RecordSource = "SELECT * FROM CLSCCB"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.ƾ֤��,3))) FROM CLSCCB WHERE CLSCCB.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "S-1"
If Not Data7.Recordset.EOF Then
PZH = "S-" + Trim(Data7.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "����ԭ����"
Data8.Recordset.Fields(1) = "�����ɱ�"
Data8.Recordset.Fields(2) = "ֱ�������ɱ�"
Data8.Recordset.Fields(3) = "�������"
Data8.Recordset.Fields(4) = ""
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(4)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = ""
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
Data4.Database.Execute "UPDATE KPD SET ���='��' WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "���ϳ��ⵥ"
Data2.Recordset.Fields(3) = "�ɱ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.ƾ֤��,3))) FROM CLSCCB WHERE CLSCCB.���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "S-1"
If Data7.Recordset.EOF Then
PZH = "S-1"
Else
PZH = "S-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Loop
Data4.Database.Execute "UPDATE KPD SET ���='��' WHERE ���� BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "���ϳ��ⵥ"
Data2.Recordset.Fields(3) = "�ɱ�ƾ֤"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
End If
End Sub

