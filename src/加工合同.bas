Attribute VB_Name = "�ӹ���ͬ"
Public Sub htht(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")


Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��ӡģ���ͬ.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_cpjf where �������='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(4, 2) = gh
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(6, 2) = Trim(DT1.Recordset.Fields(5))
Excelapp.ActiveSheet.Cells(85, 4) = DT1.Recordset.Fields(1)  ''''''������ַ
Excelapp.ActiveSheet.Cells(86, 4) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(87, 4) = DT1.Recordset.Fields(3)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct �ͻ�,����,����  from sczy_x where ����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(66, 3) = CDate(DT1.Recordset.Fields(2)) - CDate(DT1.Recordset.Fields(1))
Excelapp.ActiveSheet.Cells(66, 5) = Trim(DT1.Recordset.Fields(1))
Excelapp.ActiveSheet.Cells(66, 8) = Trim(DT1.Recordset.Fields(2))
End If


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_mlgg where �������='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(15, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(15, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(15, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(15, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(15, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(15, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(15, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(15, 10) = DT1.Recordset.Fields(9)
End If


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_flgg where �������='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(16, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(16, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(16, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(16, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(16, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(16, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(16, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(16, 10) = DT1.Recordset.Fields(9)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_qtgg where �������='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(17, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(17, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(17, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(17, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(17, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(17, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(17, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(17, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(17, 10) = DT1.Recordset.Fields(9)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct ɫ��,ɫ�ζ�  from sczy_x where ����='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 26
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(2)
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct Ʒ��,����,����,��ˮ��,Ť��,����  from sczy_x where ����='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 40
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(5)
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from htfz_cpbmyq where �������='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(52, 1) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(52, 5) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(52, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(58, 1) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(120, 2) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(145, 4) = DT1.Recordset.Fields(6)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from htfz_qybyj where �������='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(63, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(64, 2) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(82, 3) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(83, 3) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(92, 3) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(108, 3) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(109, 3) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(110, 3) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(110, 7) = DT1.Recordset.Fields(9)
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


Public Sub DXDY(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ps As Integer, sl As Single) ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ί�����.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,Ʒ��,ɫ��,����,'' as ��Լ��,ƥ��,����,'' as ����,'' as ���,����,��ע,���,����  FROM wwkpd WHERE ����='" & DH & "' and ��� between '" & xh1 & "' and '" & xh2 & "' order BY ���"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '''�ͻ�����
        Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(9))   '''����
        Excelapp.ActiveSheet.Cells(2, 12) = DH    '''����
        Excelapp.ActiveSheet.Cells(9, 6) = ps   '''ƥ��
        Excelapp.ActiveSheet.Cells(9, 7) = sl   '''����

i = 4
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)    ''''ɫ��
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(3)    ''''����
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(5)    ''''ƥ��
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(6)    ''''����
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(12)    ''''����
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(11)    '''''�ӹ����
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(10)    '''''��ע
i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "SELECT ģ�� FROM yhb WHERE �û�='" & yhm & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(11, 12) = DT1.Recordset.Fields(0)   '''�Ƶ�
End If

Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub XSHT(DT1 As Adodc, DH As String)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���ۺ�ͬ.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,����,���,Ʒ��,����+����,ɫ��,�ƻ�,����,��ע,����,����,���,�ܱ�ע,ͶȾ���,������;  FROM sczykpd WHERE ����='" & DH & "' order BY ���"
DT1.Refresh


If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(0)   '''�ͻ�����
        Excelapp.ActiveSheet.Cells(6, 7) = DT1.Recordset.Fields(1)   '''��ͬ���
        Excelapp.ActiveSheet.Cells(7, 7) = Trim(DT1.Recordset.Fields(9))    '''��ͬ����

i = 10
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(4)    ''''���
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(5)    ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)    ''''����
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(7)    ''''����
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(6) * DT1.Recordset.Fields(7), "#0.00") ''''���
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(8)    '''''��ע
i = i + 1
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

Public Sub SCTZD(DT1 As Adodc, DH As String)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim L As Integer
        Dim ym As Integer
        Dim dym As Integer

        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\����֪ͨ��.xls")
'5)���õ�2��������Ϊ�������

DT1.RecordSource = "SELECT ��� FROM sczykpd WHERE ����='" & DH & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then

L = DT1.Recordset.RecordCount
If L / 10 <> Int(L / 10) Then
ym = Int(L / 10) + 1
Else
ym = Int(L / 10)
End If
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�,����,���,Ʒ��,����,����,ɫ��,�ƻ�,ɫ��,��ע,����,����,���,�ܱ�ע,ͶȾ���,������;,�ɷ�,����,isnull(���,'') as ȷ�����  FROM sczykpd WHERE ����='" & DH & "' order BY ���"
DT1.Refresh

dym = 1
L = 1
DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(ym)   '''ҳ��
        Excelapp.ActiveSheet.Cells(3, 16) = Trim(dym)   '''�ڼ�ҳ
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   '''�ͻ�����
        Excelapp.ActiveSheet.Cells(4, 13) = Trim(DT1.Recordset.Fields(10))   '''��ͬ����
        Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(1)    '''��ͬ���
        Excelapp.ActiveSheet.Cells(5, 13) = DT1.Recordset.Fields(17)    ''��ͬҵ��
        Excelapp.ActiveSheet.Cells(20, 15) = DT1.Recordset.Fields(15)    '''������;
        Excelapp.ActiveSheet.Cells(24, 15) = Trim(DT1.Recordset.Fields(11))    '''����
        Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(13)    '''�ܱ�ע

i = 7
yj = ""
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    '''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(16)    '''�ɷ�
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(8)    ''''ɫ��
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)    ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)    ''''�ŷ�
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(5)    ''''����
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    ''''����
        Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(9)  ''''��ע
yj = Trim(yj) + Trim(i) + DT1.Recordset.Fields(18)
i = i + 1
L = L + 1
DT1.Recordset.MoveNext
If L = 11 Then
        Excelapp.ActiveSheet.Cells(19, 1) = yj    '''ȷ�����
        yj = ""
dym = dym + 1
L = 1
i = 7
        Excelapp.Sheets(dym).Activate
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(ym)   '''ҳ��
        Excelapp.ActiveSheet.Cells(3, 16) = Trim(dym)   '''�ڼ�ҳ
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   '''�ͻ�����
        Excelapp.ActiveSheet.Cells(4, 13) = Trim(DT1.Recordset.Fields(10))   '''��ͬ����
        Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(1)    '''��ͬ���
        Excelapp.ActiveSheet.Cells(5, 13) = DT1.Recordset.Fields(17)    ''��ͬҵ��
        Excelapp.ActiveSheet.Cells(20, 15) = DT1.Recordset.Fields(15)    '''������;
        Excelapp.ActiveSheet.Cells(24, 15) = Trim(DT1.Recordset.Fields(11))    '''����
        Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(13)    '''�ܱ�ע
End If
Loop
        Excelapp.ActiveSheet.Cells(19, 1) = yj    '''ȷ�����
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



