Attribute VB_Name = "��֯"
Public Sub ddlcd(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "������֯���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��֯\����\�����ƻ���.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,���,֯��,Ʒ��,Ͳ��,�ƻ�,����,����,��ɫ,������,��ע,����,����,���,ɴ��,����,ƥ�� FROM v_kpd_ddjh where ����='" & Zh & "' order by ���"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''֯��
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''�ͻ�
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''����
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''���
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4Ʒ��
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''��ɫ
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''����
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''�ƻ�

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''����
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''����
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''ƥ��
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''������
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''����
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''��ע


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT ֯��,sum(�ƻ�) FROM v_kpd_ddjh where ����='" & Zh & "' group by ֯�� order by ֯��"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT ֯��,ɴ֧,����,֯��,���,��ע,��ɫ,���� FROM sxpb where ֯��='" & DT1.Recordset.Fields(0) & "' order by ���"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''֯��
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''ɴ֧
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''��ɫ
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''����
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''����
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''֯��
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''���
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''ɴ��
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''��ע
L = L + 1
dt2.Recordset.MoveNext
Loop
End If
DT1.Recordset.MoveNext
Loop
End If


Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub ddlcdxz(DT1 As Adodc, dt2 As Adodc, Zh As String, fw As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "������֯���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��֯\����\�����ƻ���.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,���,֯��,Ʒ��,Ͳ��,�ƻ�,����,����,��ɫ,������,��ע,����,����,���,ɴ��,����,ƥ�� FROM v_kpd_ddjh where  ֯�� in(" + fw + ") order by ֯��,���"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''֯��
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''�ͻ�
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''����
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''���
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4Ʒ��
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''��ɫ
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''����
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''�ƻ�

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''����
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''����
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''ƥ��
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''������
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''����
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''��ע


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT ֯��,sum(�ƻ�) FROM v_kpd_ddjh where  ֯�� in(" + fw + ") group by ֯�� order by ֯��"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT ֯��,ɴ֧,����,֯��,���,��ע,��ɫ,���� FROM sxpb where ֯��='" & DT1.Recordset.Fields(0) & "' order by ���"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''֯��
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''ɴ֧
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''��ɫ
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''����
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''����
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''֯��
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''���
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''ɴ��
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''��ע
L = L + 1
dt2.Recordset.MoveNext
Loop
End If
DT1.Recordset.MoveNext
Loop
End If


Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub ddlcdjh(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "������֯���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��֯\����\�����ƻ���.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,���,֯��,Ʒ��,Ͳ��,�ƻ�,����,����,��ɫ,������,��ע,����,����,���,ɴ��,����,ƥ�� FROM v_kpd_ddjh_cjjt where ����='" & Zh & "' order by ���"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''֯��
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''�ͻ�
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''����
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''���
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4Ʒ��
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''��ɫ
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''����
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''�ƻ�

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''����
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''����
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''ƥ��
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''������
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''����
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''��ע


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT ֯��,sum(�ƻ�) FROM v_kpd_ddjh_cjjt where ����='" & Zh & "' group by ֯�� order by ֯��"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT ֯��,ɴ֧,����,֯��,���,��ע,��ɫ,���� FROM sxpb where ֯��='" & DT1.Recordset.Fields(0) & "' order by ���"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''֯��
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''ɴ֧
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''��ɫ
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''����
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''����
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''֯��
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''���
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''ɴ��
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''��ע
L = L + 1
dt2.Recordset.MoveNext
Loop
End If
DT1.Recordset.MoveNext
Loop
End If


Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub dbqww(DT1 As Adodc, dt2 As Adodc, DH As String, k As Long, bh As Long, cj As String, jh As String, bz As String)  ''''�ޱ���

        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "������֯���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��֯\����\tmww.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,��ͬ��,����,֯��,Ʒ��,����,����,ɴ��,����,��̨,Ͳ��,������,��ɫ FROM kpd WHERE ֯��='" & DH & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst

L = 0
TM = Mid("00000000", 1, 8 - Len(Trim(bh))) + Trim(bh)

        Excelapp.ActiveSheet.Cells(L * 20 + 2, 2) = DH    '''''֯��
        Excelapp.ActiveSheet.Cells(L * 20 + 3, 2) = DT1.Recordset.Fields(1)  '''''��ͬ��
        Excelapp.ActiveSheet.Cells(L * 20 + 3, 4) = bz  '''''������
        Excelapp.ActiveSheet.Cells(L * 20 + 4, 2) = DT1.Recordset.Fields(0)  ''''�ͻ�
        Excelapp.ActiveSheet.Cells(L * 20 + 5, 2) = DT1.Recordset.Fields(4) '''Ʒ��
        
        
        Excelapp.ActiveSheet.Cells(L * 20 + 14, 1) = "*" + TM + "J" + "*"            '''''����
        Excelapp.ActiveSheet.Cells(L * 20 + 18, 1) = "*" + TM + "J" + "*"            '''''����
        Excelapp.ActiveSheet.Cells(L * 20 + 11, 2) = DT1.Recordset.Fields(10)                         ''''''Ͳ��
        Excelapp.ActiveSheet.Cells(L * 20 + 11, 4) = DT1.Recordset.Fields(6)                         ''''''����
        
        Excelapp.ActiveSheet.Cells(L * 20 + 12, 2) = DT1.Recordset.Fields(12)                         ''''''��ɫ
        Excelapp.ActiveSheet.Cells(L * 20 + 12, 4) = cj   ''''DT1.Recordset.Fields(9)                         ''''''��̨
        
        Excelapp.ActiveSheet.Cells(L * 20 + 13, 2) = k                           ''''''ƥ��
        Excelapp.ActiveSheet.Cells(L * 20 + 13, 4) = ""                          ''''����
        Excelapp.ActiveSheet.Cells(L * 20 + 15, 4) = jh                           ''''''��̨���
        
dt2.RecordSource = "SELECT ɴ֧,����,���� FROM sxpbf WHERE ֯��='" & DH & "' and ״̬='��'"
dt2.Refresh
      
If Not dt2.Recordset.EOF Then  ''''''''''''''''''''''''''''''''''''
dt2.Recordset.MoveFirst
m = 0
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 1) = dt2.Recordset.Fields(0)  '''��ɴ
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 3) = dt2.Recordset.Fields(1)  '''����
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 4) = dt2.Recordset.Fields(2)  '''����
dt2.Recordset.MoveNext
m = m + 1
Loop
Else     '''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
DT1.RecordSource = "SELECT ɴ֧,����,���� FROM sxpb WHERE ֯��='" & DH & "'"
DT1.Refresh
      
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
m = 0
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 1) = DT1.Recordset.Fields(0)  '''��ɴ
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 3) = DT1.Recordset.Fields(1)  '''����
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 4) = DT1.Recordset.Fields(2)  '''����
DT1.Recordset.MoveNext
m = m + 1
Loop
End If
End If  ''''''''''''''''''''''''
End If

Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = False
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintOut
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Public Sub jhlcd(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "������֯���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��֯\����\������.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,���,֯��,Ʒ��,Ͳ��,�ƻ�,����,����,��ɫ,������,��ע,����,����,���,ɴ��,����,��̨,ƥ�� FROM v_kpd_ctjh where ����='" & Zh & "' order by ���"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 2
L = 2
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(16)  ''''����
Excelapp.ActiveSheet.Cells(2, 4) = DT1.Recordset.Fields(4)  ''''���
Excelapp.ActiveSheet.Cells(2, 6) = Trim(DT1.Recordset.Fields(2))  ''''֯��

Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(3)   '''Ʒ��
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(0)   '''�ͻ�
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(5)     '''''''''�ƻ�

Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)         '''ƥ��
Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(9)   ''''''������
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''���



i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT ֯��,sum(�ƻ�) FROM v_kpd_ctjh where ����='" & Zh & "' group by ֯�� order by ֯��"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT ֯��,ɴ֧,����,֯��,���,��ע,��ɫ,���� FROM sxpb where ֯��='" & DT1.Recordset.Fields(0) & "' order by ���"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(4, L) = dt2.Recordset.Fields(1)         '''ɴ֧
Excelapp.ActiveSheet.Cells(5, L) = dt2.Recordset.Fields(7)                        ''''����
Excelapp.ActiveSheet.Cells(7, L) = dt2.Recordset.Fields(2)                        ''''����
L = L + 1
dt2.Recordset.MoveNext
Loop
End If
DT1.Recordset.MoveNext
Loop
End If


Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


