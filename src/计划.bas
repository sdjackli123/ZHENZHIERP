Attribute VB_Name = "�ƻ�"
Public Sub jhbOutadodcToExcel(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\jhb.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j - 1)
                      
                          Next j
               
         If i >= 1 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
Excelapp.ActiveSheet.Cells(1, 1) = BT
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

Public Sub jh3(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ӡ�����̵�����.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select max(ӡ������) as zl from v_yhjh where ӡ������='" & gh & "' and ��� is not null and len(���)>0"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from v_yhjh where ӡ������='" & gh & "' And ӡ������ = '" & a & "' and ��� is not null and len(���)>0"
DT1.Refresh
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '�ͻ�
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))   '''���
Excelapp.ActiveSheet.Cells(2, 10) = Trim(DT1.Recordset.Fields(23))   ''''����
Excelapp.ActiveSheet.Cells(2, 13) = Trim(DT1.Recordset.Fields(3))      ''''''����
'''Excelapp.ActiveSheet.Cells(3, 2) = dt1.Recordset.Fields(13)        ''''
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(5)     '''Ʒ��
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(4)    ''ɫ��
Excelapp.ActiveSheet.Cells(5, 3) = Trim(gh)                           ''ӡ������
Excelapp.ActiveSheet.Cells(3, 4) = DT1.Recordset.Fields(24)    'ӡ�����
Excelapp.ActiveSheet.Cells(5, 9) = "*" + Trim(DT1.Recordset.Fields(3)) + "J*"            ''' Ⱦɫ��������

DT1.RecordSource = "select round(sum(ӡ������),2),sum(ӡ��ƥ��) from v_yhjh where ӡ������='" & gh & "' and ��� is not null and len(���)>0"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Excelapp.ActiveSheet.Cells(4 + 22, 9) = DT1.Recordset.Fields(0)
End If

DT1.RecordSource = "select * from v_yhjh where ӡ������='" & gh & "' and ��� is not null and len(���)>0"
DT1.Refresh
i = 0
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(10 + i * 1, 1) = DT1.Recordset.Fields(5)    '''Ʒ��
Excelapp.ActiveSheet.Cells(10 + i * 1, 4) = DT1.Recordset.Fields(21)   ''
Excelapp.ActiveSheet.Cells(10 + i * 1, 6) = DT1.Recordset.Fields(20)
Excelapp.ActiveSheet.Cells(10 + i * 1, 7) = DT1.Recordset.Fields(19)  ''��λ
Excelapp.ActiveSheet.Cells(10 + i * 1, 8) = DT1.Recordset.Fields(17) '''''''''���
Excelapp.ActiveSheet.Cells(10 + i * 1, 9) = DT1.Recordset.Fields(22)  '''��ע
Excelapp.ActiveSheet.Cells(10 + i * 1, 13) = DT1.Recordset.Fields(18) '''''ͼ��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
i = i + 1
DT1.Recordset.MoveNext
Loop


Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub

Public Sub jh33(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ӡ�����̵�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select distinct ӡ������,�ͻ�����,CONVERT(varchar,�ƻ�����, 23),���,ӡ����� from v_yhjh where ӡ������='" & gh & "' and ��� is not null and len(���)>0 order by ӡ������"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(2, 14) = Trim(gh)
Excelapp.ActiveSheet.Cells(2, 16) = DT1.Recordset.Fields(4)

i = 4
Do While Not DT1.Recordset.EOF

dt2.RecordSource = "select * from v_yhjh where ӡ������='" & DT1.Recordset.Fields(0) & "' and ��� is not null and len(���)>0"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(5)    '''Ʒ��
Excelapp.ActiveSheet.Cells(i, 4) = dt2.Recordset.Fields(3)     ''''����
Excelapp.ActiveSheet.Cells(i, 5) = dt2.Recordset.Fields(4)     ''ɫ��
Excelapp.ActiveSheet.Cells(i, 6) = Trim(dt2.Recordset.Fields(21))  'ƥ��
Excelapp.ActiveSheet.Cells(i, 7) = Trim(dt2.Recordset.Fields(20))  ''����
Excelapp.ActiveSheet.Cells(i, 8) = dt2.Recordset.Fields(19)  ''��λ
Excelapp.ActiveSheet.Cells(i, 9) = dt2.Recordset.Fields(16)  ''���
Excelapp.ActiveSheet.Cells(i, 10) = dt2.Recordset.Fields(22)    '''��ע
Excelapp.ActiveSheet.Cells(i, 14) = dt2.Recordset.Fields(18) '''''ͼ��
Excelapp.ActiveSheet.Cells(i, 16) = dt2.Recordset.Fields(17)  '''''''''���
i = i + 1
dt2.Recordset.MoveNext
Loop

DT1.Recordset.MoveNext
Loop
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
Excelapp.Quit
Set Excelapp = Nothing
End Sub


Public Sub pcjh(DT1 As Adodc, dt2 As Adodc, sql1 As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\Ⱦɫ�ƻ�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.RecordSource = "SELECT ��̨��� FROM CT ORDER BY ip"
DT1.Refresh


If Not DT1.Recordset.EOF Then
i = 4
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
dt2.RecordSource = "SELECT top 6 * FROM v_kpdb where  (" + sql1 + ") and ��̨='" & DT1.Recordset.Fields(0) & "' ORDER BY �Ų����"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
L = 1
Do While Not dt2.Recordset.EOF
'Excelapp.ActiveSheet.Cells(i, 2) = dt2.Recordset.Fields(0)    '''Ʒ��
Excelapp.ActiveSheet.Cells(i, 3) = dt2.Recordset.Fields(1)     ''''�Ų�ʱ��
Excelapp.ActiveSheet.Cells(i, 4) = dt2.Recordset.Fields(2)     ''�Ų����
Excelapp.ActiveSheet.Cells(i, 5) = dt2.Recordset.Fields(3)  '�ͻ�����
Excelapp.ActiveSheet.Cells(i, 6) = dt2.Recordset.Fields(4)  ''Ʒ��
Excelapp.ActiveSheet.Cells(i, 7) = dt2.Recordset.Fields(5)  ''ɫ��
Excelapp.ActiveSheet.Cells(i, 8) = dt2.Recordset.Fields(6)  ''��ɫ
Excelapp.ActiveSheet.Cells(i, 9) = dt2.Recordset.Fields(7)    '''����
Excelapp.ActiveSheet.Cells(i, 10) = dt2.Recordset.Fields(8) ''''����
Excelapp.ActiveSheet.Cells(i, 11) = dt2.Recordset.Fields(9)  '''''''''�Ų���ע
Excelapp.ActiveSheet.Cells(i, 12) = dt2.Recordset.Fields(9)  '''''''''����
i = i + 1
dt2.Recordset.MoveNext
L = L + 1
Loop
i = i + 7 - L
Else
i = i + 6
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

Public Sub pgk(DT1 As Adodc, gh As String, selectedPrinter As String)
    Dim Excelapp As Object  ' ���� Excel Ӧ�ó������
    On Error Resume Next    ' ���ִ���ʱ����ִ����һ��

    Set Excelapp = CreateObject("Excel.Application")  ' ���� Excel Ӧ�ó������

    If Excelapp Is Nothing Then   ' ��� Excel Ӧ�ó������δ�����ɹ�
        MsgBox "Excel is not installed on this machine."   ' ��ʾ������Ϣ
        Exit Sub   ' �˳��ӳ���
    End If

    On Error GoTo Ert   ' ���ִ���ʱת�� Ert ��ǩ

    Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"   ' ���� Excel ���ڱ���
    Excelapp.SheetsInNewWorkbook = 1   ' �����¹������еĹ�������Ϊ 1
    Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�Ÿ׿�.xls")   ' ���Ѵ��ڵĹ�����
    Excelapp.Sheets(1).Activate   ' �����һ��������

    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"   ' �������ݿ������ַ���
    DT1.RecordSource = "select �ͻ�����,����,Ʒ��,��ǩ,ɫ��,ƥ��,����,dr,����,ɫ��,����Ҫ��,��ע from kpd where ����='" & gh & "' "   ' ���ò�ѯ���
    DT1.Refresh   ' ˢ������

    If Not DT1.Recordset.EOF Then   ' �����¼����Ϊ��
        DT1.Recordset.MoveFirst   ' �ƶ�����¼���ĵ�һ����¼
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   ' ���ͻ�������д����Ԫ��
        Excelapp.ActiveSheet.Cells(3, 4) = Trim(DT1.Recordset.Fields(1))   ' ��������д����Ԫ��
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(2)   ' ��Ʒ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)   ' ����ɫ��д����Ԫ��
        Excelapp.ActiveSheet.Cells(5, 4) = DT1.Recordset.Fields(9)   ' ��ɫ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(5)   ' ��ƥ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(6, 4) = DT1.Recordset.Fields(6)   ' ��������д����Ԫ��
    End If

    Excelapp.ActiveWindow.Zoom = 100   ' ���ô������ű���Ϊ 100%
    ' ExcelApp.Visible = True  ' ע�͵����� Excel Ӧ�ó���ɼ��Ĵ���
    Excelapp.DisplayAlerts = False   ' ������ʾ����

    ' �л����û�ѡ��Ĵ�ӡ��
    If selectedPrinter <> "" Then
        TrySetActivePrinter Excelapp, selectedPrinter
    End If

    ' ��ӡ������
    Excelapp.ActiveSheet.PrintOut Copies:=1, Preview:=False, PrintToFile:=False, Collate:=True

Cleanup:
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
    Exit Sub

Ert:   ' �������ǩ
    MsgBox "An error occurred: " & Err.Description   ' ��ʾ������Ϣ
    Excelapp.Quit   ' �˳� Excel Ӧ�ó���
    Set Excelapp = Nothing   ' �ͷ� Excel Ӧ�ó������
End Sub
Public Sub pgk1(DT1 As Adodc, gh As String)
    Dim Excelapp As Object  ' ���� Excel Ӧ�ó������
    On Error Resume Next    ' ���ִ���ʱ����ִ����һ��

    Set Excelapp = CreateObject("Excel.Application")  ' ���� Excel Ӧ�ó������

    If Excelapp Is Nothing Then   ' ��� Excel Ӧ�ó������δ�����ɹ�
        MsgBox "Excel is not installed on this machine."   ' ��ʾ������Ϣ
        Exit Sub   ' �˳��ӳ���
    End If

    On Error GoTo Ert   ' ���ִ���ʱת�� Ert ��ǩ

    Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"   ' ���� Excel ���ڱ���
    Excelapp.SheetsInNewWorkbook = 1   ' �����¹������еĹ�������Ϊ 1
    Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�Ÿ׿�.xls")   ' ���Ѵ��ڵĹ�����
    Excelapp.Sheets(1).Activate   ' �����һ��������

    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"   ' �������ݿ������ַ���
    DT1.RecordSource = "select �ͻ�����,����,Ʒ��,��ǩ,ɫ��,ƥ��,����,dr,����,ɫ��,����Ҫ��,��ע from kpd where ����='" & gh & "' "   ' ���ò�ѯ���
    DT1.Refresh   ' ˢ������

    If Not DT1.Recordset.EOF Then   ' �����¼����Ϊ��
        DT1.Recordset.MoveFirst   ' �ƶ�����¼���ĵ�һ����¼
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   ' ���ͻ�������д����Ԫ��
        Excelapp.ActiveSheet.Cells(3, 4) = Trim(DT1.Recordset.Fields(1))   ' ��������д����Ԫ��
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(2)   ' ��Ʒ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)   ' ����ɫ��д����Ԫ��
        Excelapp.ActiveSheet.Cells(5, 4) = DT1.Recordset.Fields(9)   ' ��ɫ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(5)   ' ��ƥ����д����Ԫ��
        Excelapp.ActiveSheet.Cells(6, 4) = DT1.Recordset.Fields(6)   ' ��������д����Ԫ��
    End If

    Excelapp.ActiveWindow.Zoom = 100   ' ���ô������ű���Ϊ 100%
    ' Excelapp.Visible = True  ' ע�͵����� Excel Ӧ�ó���ɼ��Ĵ���
    Excelapp.DisplayAlerts = False   ' ������ʾ����

    Excelapp.ActiveSheet.PrintOut   ' ֱ�Ӵ�ӡ��ǰ������

    Set Excelapp = Nothing   ' �ͷ� Excel Ӧ�ó������
    Exit Sub   ' �˳��ӳ���

Ert:   ' �������ǩ
    MsgBox "An error occurred: " & Err.Description   ' ��ʾ������Ϣ
    Excelapp.Quit   ' �˳� Excel Ӧ�ó���
    Set Excelapp = Nothing   ' �ͷ� Excel Ӧ�ó������
End Sub

Private Sub TrySetActivePrinter(ByRef Excelapp As Object, ByVal PrinterName As String)
    On Error Resume Next
    Dim CurrentPrinter As String
    CurrentPrinter = Excelapp.ActivePrinter
    Excelapp.ActivePrinter = PrinterName
    If Err.Number <> 0 Then
        ' ���Ը��Ӷ˿�����
        Excelapp.ActivePrinter = PrinterName & " on " & Split(PrinterName, " (")(1) ' ��ȡ�����Ӷ˿�����
        If Err.Number = 0 Then
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub pcmx(Flex As VSFlexGrid, BT As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\pcmx.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k + 1

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT


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

