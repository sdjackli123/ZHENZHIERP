Attribute VB_Name = "ƾ֤����"

Public Sub FKHZXJ()   '''''''''�������--�ֽ�
On Error Resume Next

Adodc8.RecordSource = "SELECT * FROM CLFKPZ"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'2-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "2-1"
If Not Adodc7.Recordset.EOF Then
PZH = "2-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM JGZCX1 where val(�����ָ���)>0"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = "������"     '''''ժҪ
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)      ''''�跽���˿�Ŀ
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)         ''''�跽��ϸ��Ŀ
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(4)
End If
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)     '''''''''''�������˿�Ŀ
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)       '''������ϸ��Ŀ
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(0)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(6)     ''''''''�������
Adodc8.Recordset.Fields(6) = PZH                           '''''''''''''ƾ֤��
Adodc8.Recordset.Fields(7) = DTPicker3.Value               '''''''��������
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""                           ''''''''''����
Adodc8.Recordset.Fields(10) = ""                          ''''''''''''����
Adodc8.Recordset.Fields(11) = adodcCombo3.Text       ''''''''''''''�Ƶ�
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "��"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ𸶿�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�ֽ𸶿"
Adodc2.Recordset.Fields(3) = "����ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'2-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "2-1"
If Not Adodc7.Recordset.EOF Then
PZH = "2-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ𸶿�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�ֽ𸶿"
Adodc2.Recordset.Fields(3) = "����ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
End If
End Sub
''''''''''''''
Public Sub FKHZYH()   '''''''''�������---���д��
On Error Resume Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'4-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "4-1"
If Not Adodc7.Recordset.EOF Then
PZH = "4-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �������<>'0' AND �跽���='0' AND INSTR(���,'���д��')>0 AND (���<>'��' OR ���=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(4)
End If
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(0)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(6)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "��"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "���и��"
Adodc2.Recordset.Fields(3) = "����ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLFKPZ WHERE INSTR(ƾ֤��,'4-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "4-1"
If Not Adodc7.Recordset.EOF Then
PZH = "4-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "���и��"
Adodc2.Recordset.Fields(3) = "����ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���ƾ֤")
End If
End Sub

Public Sub SKHZXJ()    ''''''''�տ����----�ֽ�
On Error Resume Next
If adodcCombo3.Text = "" Then
MsgBox ("��ѡ�񸴺�Ա")
Exit Sub
End If
Adodc8.RecordSource = "SELECT * FROM CLSKPZ"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'1-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "1-1"
If Not Adodc7.Recordset.EOF Then
PZH = "1-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �跽���<>'0' AND �������='0' AND INSTR(���,'�ֽ�')>0 AND (���<>'��' OR ���=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(0)
End If
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(4)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(5)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "��"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ��տ�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�ֽ��տ"
Adodc2.Recordset.Fields(3) = "�տ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'1-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "1-1"
If Not Adodc7.Recordset.EOF Then
PZH = "1-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�ֽ��տ"
Adodc2.Recordset.Fields(3) = "�տ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ֽ��տ�ƾ֤")
End If
End Sub

Public Sub SKHZYH()    ''''''''�տ����----���д��
On Error Resume Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'3-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "3-1"
If Not Adodc7.Recordset.EOF Then
PZH = "3-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Adodc9.RecordSource = "SELECT * FROM TZJZMX WHERE ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime) AND (���ȷ��<>'��' OR ���ȷ��=NULL) AND �跽���<>'0' AND �������='0' AND INSTR(���,'���д��')>0 AND (���<>'��' OR ���=NULL)"
Adodc9.Refresh
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = Adodc9.Recordset.Fields(3)
If InStr(Adodc9.Recordset.Fields(0), "-") > 0 Then
Adodc8.Recordset.Fields(1) = Left(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") - 1)
Adodc8.Recordset.Fields(2) = Mid(Adodc9.Recordset.Fields(0), InStr(Adodc9.Recordset.Fields(0), "-") + 1)
Else
Adodc8.Recordset.Fields(1) = Adodc9.Recordset.Fields(0)
End If
If InStr(Adodc9.Recordset.Fields(4), "-") > 0 Then
Adodc8.Recordset.Fields(3) = Left(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") - 1)
Adodc8.Recordset.Fields(4) = Mid(Adodc9.Recordset.Fields(4), InStr(Adodc9.Recordset.Fields(4), "-") + 1)
Else
Adodc8.Recordset.Fields(3) = Adodc9.Recordset.Fields(4)
End If
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(5)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = Adodc9.Recordset.Fields(2)
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.Edit    '''''''''''''''''''''''''''''''
Adodc9.Recordset.Fields(12) = "��"
Adodc9.Recordset.Update    ''''''''''''''''''''''''''''''
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���տ�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�����տ"
Adodc2.Recordset.Fields(3) = "�տ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
Adodc7.RecordSource = "SELECT MAX(VAL(MID(ƾ֤��,3))) FROM CLSKPZ WHERE INSTR(ƾ֤��,'3-')>0 AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "3-1"
If Not Adodc7.Recordset.EOF Then
PZH = "3-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "�����տ"
Adodc2.Recordset.Fields(3) = "�տ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "���д���տ�ƾ֤")
End If
End Sub




Public Sub CLCKpz()
On Error Resume Next
Adodc8.adodcbase.Execute "DELETE * FROM CLSCHZ"
lo = "d:\���ݿ�\bfrz\" + ljb + "\cw.mdb"       '''''''''''''''''''''''����
Adodc3.adodcbase.Execute "INSERT INTO CLSCHZ(���) IN'" & lo & "' SELECT FORMAT(SUM(�ϼƽ��),'#0.00') AS ��� FROM KPD WHERE (���<>'��' OR ���=NULL) AND ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc8.adodcbase.Execute "DELETE * FROM CLSCHZ WHERE ���=NULL"
Adodc9.RecordSource = "SELECT * FROM CLSCHZ"
Adodc9.Refresh
Adodc8.RecordSource = "SELECT * FROM CLSCCB"
Adodc8.Refresh
Adodc7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.ƾ֤��,3))) FROM CLSCCB WHERE CLSCCB.���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "S-1"
If Not Adodc7.Recordset.EOF Then
PZH = "S-" + Trim(Adodc7.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If
If Adodc9.Recordset.EOF Then
Exit Sub
Else
Adodc9.Recordset.MoveFirst
KLLLL = 1
Do While Not Adodc9.Recordset.EOF
For i = 1 To 5
Adodc8.Recordset.AddNew
Adodc8.Recordset.Fields(0) = "����ԭ����"
Adodc8.Recordset.Fields(1) = "�����ɱ�"
Adodc8.Recordset.Fields(2) = "ֱ�������ɱ�"
Adodc8.Recordset.Fields(3) = "�������"
Adodc8.Recordset.Fields(4) = ""
Adodc8.Recordset.Fields(5) = Adodc9.Recordset.Fields(4)
Adodc8.Recordset.Fields(6) = PZH
Adodc8.Recordset.Fields(7) = DTPicker3.Value
Adodc8.Recordset.Fields(8) = ""
Adodc8.Recordset.Fields(9) = ""
Adodc8.Recordset.Fields(10) = ""
Adodc8.Recordset.Fields(11) = adodcCombo3.Text
Adodc8.Recordset.Update
Adodc9.Recordset.MoveNext
If Adodc9.Recordset.EOF Then
Adodc4.adodcbase.Execute "UPDATE KPD SET ���='��' WHERE ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "���ϳ��ⵥ"
Adodc2.Recordset.Fields(3) = "�ɱ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Adodc7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.ƾ֤��,3))) FROM CLSCCB WHERE CLSCCB.���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
Adodc7.Refresh
PZH = "S-1"
If Adodc7.Recordset.EOF Then
PZH = "S-1"
Else
PZH = "S-" + Trim(Adodc7.Recordset.Fields(0) + 1)
End If
Loop
Adodc4.adodcbase.Execute "UPDATE KPD SET ���='��' WHERE ���� BETWEEN cast('" & Text1.Text & "' as datetime) AND cast('" & Text2.Text & "' as datetime)"
MsgBox ("ת�˳ɹ���" + "����" + Str(KLLLL) + "�ɱ�ƾ֤")
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = DTPicker3.Value
Adodc2.Recordset.Fields(1) = Text1.Text
Adodc2.Recordset.Fields(2) = "���ϳ��ⵥ"
Adodc2.Recordset.Fields(3) = "�ɱ�ƾ֤"
Adodc2.Recordset.Fields(4) = Str(KLLLL)
Adodc2.Recordset.Update
Adodc2.Refresh
End If
End Sub

