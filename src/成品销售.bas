Attribute VB_Name = "��Ʒ����"
Public Sub dzmx(Flex As VSFlexGrid, fd1, fd2, BT)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\dzmx.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k + 1

                          For j = 1 To .Cols

                              
                              DoEvents
                              Excelapp.ActiveSheet.Cells(i + 2, j) = "'" & .TextMatrix(i - 1, j)
                              If i >= 2 And (j = (fd1 - 1) Or j = (fd2 - 1)) Then
                              Excelapp.ActiveSheet.Cells(i + 2, j) = Val(Excelapp.ActiveSheet.Cells(i + 2, j))
                              End If
                          Next j
               
         If i >= 2 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i + 2, fd1 - 1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i + 2, fd2 - 1)) + Q2
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(2, 2) = BT + "  �ͻ����˵�"

Excelapp.ActiveSheet.Cells(i + 1, 1) = "�ϼ�"
Excelapp.ActiveSheet.Cells(i + 1, fd1 - 1) = Q1
Excelapp.ActiveSheet.Cells(i + 1, fd2 - 1) = Q2

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



Public Sub OutadodcToExcelBC(Flex As VSFlexGrid, FD, BT)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
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
               
         If i >= 3 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
Excelapp.ActiveSheet.Cells(i, FD) = Q

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

Public Sub CPCKOutadodcToExcel(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ps As Integer, sl As Single) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ����.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ�,isnull(����,0)  FROM v_jgmx WHERE ����='" & DH & "' and ˳��� between '" & xh1 & "' and '" & xh2 & "' order BY ˳���"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '''�ͻ�����
        Excelapp.ActiveSheet.Cells(2, 6) = Trim(DT1.Recordset.Fields(9))   '''����
        Excelapp.ActiveSheet.Cells(2, 10) = DH    '''����
        Excelapp.ActiveSheet.Cells(9, 7) = ps   '''ƥ��
        Excelapp.ActiveSheet.Cells(9, 8) = sl   '''����

i = 4
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    ''''����
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)    ''''���
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(2)    ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(5)    ''''ƥ��
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(6)    ''''����
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(13)    '''''����
        Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(10)   ''''��ע
i = i + 1
DT1.Recordset.MoveNext
Loop


DT1.RecordSource = "SELECT round(sum(isnull(����,0)),2)  FROM v_jgmx WHERE ����='" & DH & "'"
DT1.Refresh
If Not IsNull(DT1.Recordset.Fields(0)) Then
        Excelapp.ActiveSheet.Cells(9, 9) = DT1.Recordset.Fields(0)   '''��������
End If

DT1.RecordSource = "SELECT ģ��  FROM yhb WHERE �û�='" & yhm & "'"
DT1.Refresh
If Not IsNull(DT1.Recordset.Fields(0)) Then
    Excelapp.ActiveSheet.Cells(11, 10) = Trim(DT1.Recordset.Fields(0))  '''�û�ģ��
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



Public Sub CPCKTZD(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ����.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT �ͻ�ȫ��,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ� as ɫ��,����,isnull(����,0),����,�ƻ���,isnull(����,0),����,'' as �µ�����,����  FROM v_jgmx WHERE ����='" & DH & "' and ˳��� between '" & xh1 & "' and '" & xh2 & "' order BY ˳���"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''�ͻ�
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''����
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''����
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''����
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''���
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''�׺�
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''ƥ��
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''ë������
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''��������
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''����
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''���
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''��ע
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'dt3.RecordSource = "SELECT SUM(ISNULL(ƥ��,0)),SUM(ISNULL(����,0)),SUM(ISNULL(����,0)),SUM(ISNULL(���,0))  FROM v_jgmx WHERE ����='" & DH & "' and (�ӹ����='��Ʒ��' or �ӹ����='Ⱦɫ��' or �ӹ����='���ͷ�' or �ӹ����='���շ�' or �ӹ����='��ӡ��' or �ӹ����='ֻĥë')"
'dt3.Refresh
'If Not dt3.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(10, 1) = "ҳ" + Trim(ym)  ''''ҳ��
'        Excelapp.ActiveSheet.Cells(10, 3) = "��ҳ" + Trim(ys)  ''''ҳ��
'        Excelapp.ActiveSheet.Cells(10, 4) = "�ϼ�"    ''''Ʒ��
'        Excelapp.ActiveSheet.Cells(10, 8) = Format(dt3.Recordset.Fields(0), "#0.0")   ''''ƥ��
'        Excelapp.ActiveSheet.Cells(10, 9) = Format(dt3.Recordset.Fields(1), "#0.00")   ''''ë������
'        Excelapp.ActiveSheet.Cells(10, 10) = Format(dt3.Recordset.Fields(2), "#0.00")   ''''��������
'End If

'dt3.RecordSource = "SELECT SUM(ISNULL(���,0))  FROM v_jgmx WHERE ����='" & DH & "'"
'dt3.Refresh
'If Not dt3.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(10, 11) = Format(dt3.Recordset.Fields(0), "#0.00") ''''���
'End If

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
Public Sub CPCKTZDGP(DT1 As Adodc, dt3 As Adodc, DT4 As Adodc, dt5 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer)

    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    Dim x   As Integer
    On Error GoTo Ert

    ' ����ExcelӦ�ó���ʵ��
    Dim Excelapp   As Excel.Application
    Set Excelapp = New Excel.Application

    On Error Resume Next

    ' ��ָ���Ĺ�����
    Excelapp.SheetsInNewWorkbook = 1
    Excelapp.Caption = "����Ⱦ�����֮��ӡ"
    Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ�������.xls")

    ' �����һ��������
    Excelapp.Sheets(1).Activate

    ' ��������Դ
    DT1.RecordSource = "SELECT �ͻ�ȫ��,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ� as ɫ��,����,isnull(����,0),����,�ƻ���,isnull(����,0),����,'' as �µ�����,����,����Ҫ��,���ϵ�λ,ҵ��  FROM v_jgmx WHERE ����='" & DH & "' and ˳��� between '" & xh1 & "' and '" & xh2 & "' order BY ˳���"
    DT1.Refresh
    dt5.RecordSource = "SELECT * FROM yskzcx where �ͻ�= '" & DT1.Recordset.Fields(0).value & "'"
    dt5.Refresh
    DT4.RecordSource = "SELECT round(sum(isnull(Ƿ��,0)),2) as �ϼ�Ƿ�� FROM jgzcx where �ͻ�= '" & DT1.Recordset.Fields(0).value & "'"
    DT4.Refresh

    ' ���Excel����
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   ' �ͻ�
        Excelapp.ActiveSheet.Cells(3, 13) = Trim(DT1.Recordset.Fields(9))  ' ����
        Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(18)       ' ����
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)       ' ����
        Excelapp.ActiveSheet.Cells(3, 7) = DT4.Recordset.Fields(0) & "Ԫ"  ' �ۼ�Ƿ��
        
        i = 5
        Do While Not DT1.Recordset.EOF
            Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ' Ʒ��
            Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ' �׺�
            Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)    ' ��ɫ
            Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ' ƥ��
            Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ' ë������
            Excelapp.ActiveSheet.Cells(i, 9) = Val(DT1.Recordset.Fields(21)) ' ����
            Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(7)   ' ����
            Excelapp.ActiveSheet.Cells(i, 11) = Val(DT1.Recordset.Fields(8)) ' ���
            Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(22)  ' ����
            Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(10)  ' ��ע
            Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(13)  ' ҵ��
            i = i + 1
            DT1.Recordset.MoveNext
        Loop
    End If

    ' ��ʾExcel�������û����б༭
    Excelapp.Visible = True
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.DisplayAlerts = False
    
    ' ������ӡԤ��
    Excelapp.ActiveSheet.PrintPreview
    
    ' ��ӡ���˳�
    Excelapp.Quit
    Set Excelapp = Nothing
    
    Exit Sub

Ert:
    ' ������
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
End Sub


Public Sub CPCKTZDF(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ����.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT �ͻ�ȫ��,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ� as ɫ��,����,isnull(����,0),����,�ƻ���,isnull(����,0),����,'' as �µ�����,����  FROM v_jgmx WHERE ����='" & DH & "' and ˳��� between '" & xh1 & "' and '" & xh2 & "'  order BY ˳���"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''�ͻ�
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''����
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''����
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''����
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''���
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''�׺�
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''ƥ��
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''ë������
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''��������
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''����
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''���
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''��ע
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If
'        Excelapp.ActiveSheet.Cells(10, 1) = "ҳ" + Trim(ym)  ''''ҳ��
'        Excelapp.ActiveSheet.Cells(10, 3) = "��ҳ" + Trim(ys)  ''''ҳ��

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

Public Sub CPCKTZDFGP(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ�������.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT �ͻ�ȫ��,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ� as ɫ��,����,isnull(����,0),����,�ƻ���,isnull(����,0),����,'' as �µ�����,����  FROM v_jgmx WHERE ����='" & DH & "' and ˳��� between '" & xh1 & "' and '" & xh2 & "'  order BY ˳���"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''�ͻ�
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''����
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''����
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''����
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''���
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''�׺�
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''ƥ��
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''ë������
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''��������
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''����
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''���
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''��ע
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If
'        Excelapp.ActiveSheet.Cells(10, 1) = "ҳ" + Trim(ym)  ''''ҳ��
'        Excelapp.ActiveSheet.Cells(10, 3) = "��ҳ" + Trim(ys)  ''''ҳ��

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

Public Sub CLRKOutadodcToExcel(DT1 As Adodc, DH As String)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\CLRK.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(2).Activate
DT1.RecordSource = "SELECT ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,����,�ϼƽ��,����  FROM CLGL WHERE ���ݺ�='" & DH & "' order BY ���"
DT1.Refresh
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 5) = DH
        Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(9)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(8)
        
i = i + 1
DT1.Recordset.MoveNext
Loop
DT1.RecordSource = "SELECT *  FROM CKGL WHERE ���ݺ�='" & DH & "' order BY ���"
DT1.Refresh


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

Public Sub CPCKQKT(DT1 As Adodc, DH As String, je As Single) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\Ƿ����.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        
DT1.RecordSource = "SELECT �ͻ�����,Ʒ��,��ɫ,����,��Լ��,ƥ��,����,����,���,����,��ע,�ӹ����,��Ʊ�ѿ�  FROM v_jgmx WHERE ����='" & DH & "'"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(4, 6) = DH  '''����
        Excelapp.ActiveSheet.Cells(6, 6) = Format(je, "#0.00")  '''���
         Excelapp.ActiveSheet.Cells(9, 6) = DT1.Recordset.Fields(0) '''�ͻ�
       Excelapp.ActiveSheet.Cells(14, 6) = Trim(DT1.Recordset.Fields(9))  '''����

Excelapp.ActiveWindow.Zoom = 100


        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

Excelapp.Quit
Set Excelapp = Nothing
End Sub


Public Sub fhdmd(DT1 As Adodc, gh As String, DH As String, dj As String)   ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�ȫ��,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����  FROM v_bmd WHERE ����='" & gh & "' and �׺�='" & DH & "' and ����='" & dj & "' order by ƥ��"
DT1.Refresh

If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 7) = Trim(DT1.Recordset.Fields(10))
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(7)
i = 1
L = 0
Do While Not DT1.Recordset.EOF
If Int(i / 19) = i / 19 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select  ����,count(ƥ��) as ƥ�� from v_bmd where ����='" & gh & "' and �׺�='" & DH & "' and ����='" & dj & "' group by ����"
DT1.Refresh

mpzl = 0
mpps = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
DT1.Recordset.MoveNext
Loop
End If

'Excelapp.ActiveSheet.Cells(32, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps


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

Public Sub CPCKSH(DT1 As Adodc, DH As String)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Ʒ�������OK.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT �ͻ�,Ʒ��,��ɫ,����,��Լ�� as �ͻ���ͬ��,ƥ��,����,����,���,����,isnull(����,0),�ƻ���,����+'*'+����,isnull(����,0),����,����  FROM v_jgmx WHERE ����='" & DH & "'  order BY ˳���"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   '''������λ
        Excelapp.ActiveSheet.Cells(2, 5) = Date   '''����
        Excelapp.ActiveSheet.Cells(2, 9) = DH    '''����
        Excelapp.ActiveSheet.Cells(17, 2) = DT1.Recordset.Fields(14)   '''ҵ��
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(12)  ''''����
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)    ''''��ɫ
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(5)    ''''ƥ��
        If Val(DT1.Recordset.Fields(13)) <> 0 Then
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(13)   ''''����
        Else
        Excelapp.ActiveSheet.Cells(i, 5) = ""  ''''����
        End If
        If DT1.Recordset.Fields(20) = "ë��" Then
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(6), "#0.00")   ''''ë������
        Else
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(10), "#0.00")   ''''��������
        End If
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(7)    '''''����
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(8))    '''''���
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(11) ''''�ƻ���
        Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(4)    '''''�ͻ���ͬ��
i = i + 1
DT1.Recordset.MoveNext
Loop


DT1.RecordSource = "SELECT SUM(ISNULL(ƥ��,0)),SUM(ISNULL(����,0)),SUM(ISNULL(����,0)),SUM(ISNULL(���,0)),round(sum(isnull(����,0)),1)  FROM v_jgmx WHERE ����='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(14, 1) = "�ϼ�"    ''''Ʒ��
        Excelapp.ActiveSheet.Cells(14, 4) = Format(DT1.Recordset.Fields(0), "#0.0")   ''''ƥ��
        Excelapp.ActiveSheet.Cells(14, 5) = Format(DT1.Recordset.Fields(4), "#0.0")   ''''ƥ��
        Excelapp.ActiveSheet.Cells(14, 6) = Format(DT1.Recordset.Fields(2), "#0.00")   ''''ë������
        Excelapp.ActiveSheet.Cells(14, 8) = Format(DT1.Recordset.Fields(3), "#0.00") ''''���
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

Public Sub rkdmd(DT1 As Adodc, gh As String, dj As String, xh As String, DH As String, xh1 As Integer)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\rkmd.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,�ɷ�,ƥ��,����,��ͬ����,�µ�����,����,����  FROM v_bmd WHERE ����='" & gh & "' and ��ⵥ��='" & dj & "' and ������='" & xh1 & "'  and ����='" & DH & "' and ���='" & xh & "' order by ƥ��"
DT1.Refresh

If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(14)      '''����
        Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(11)    '''��ͬ����
        Excelapp.ActiveSheet.Cells(5, 3) = Trim(DT1.Recordset.Fields(10))  '''�������
        Excelapp.ActiveSheet.Cells(5, 9) = Trim(DT1.Recordset.Fields(12))  '''�µ�����
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0)         '''�ͻ�
        Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(3)         '''Ʒ��
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2)         '''���
        Excelapp.ActiveSheet.Cells(8, 5) = DT1.Recordset.Fields(1)         ''����
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(8)    ''''�ɷ�
        Excelapp.ActiveSheet.Cells(8, 11) = DT1.Recordset.Fields(5)    '''ɫ��
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4)    '''����
        Excelapp.ActiveSheet.Cells(10, 5) = DT1.Recordset.Fields(7)    '''����
i = 1
L = 0
mpps = 0
mpzl = 0
Do While Not DT1.Recordset.EOF
If Int(i / 19) = i / 19 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = Val(DT1.Recordset.Fields(13))
mpps = mpps + 1
mpzl = mpzl + Val(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

Excelapp.ActiveSheet.Cells(10, 8) = mpps

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

