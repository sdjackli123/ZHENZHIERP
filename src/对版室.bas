Attribute VB_Name = "�԰���"
Public Sub gd(bg As VSFlexGrid)
  Dim Row, col As Integer                        '���������������ڽ��ձ���������
  With bg
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = "x"       '������е�ÿһ����Ԫ��ֵΪRow+Col
    Next Row
  End With
End Sub
Public Sub shbj(DT1 As Adodc, dt2 As Adodc, DH As String)
    ' ��������
    Dim i As Integer
    Dim Excelapp As Excel.Application
    Dim wb As Excel.Workbook
    Dim sh As Excel.Worksheet

    ' ����ExcelӦ��ʵ��
    Set Excelapp = New Excel.Application
    ' �¹������а���10��������
    Excelapp.SheetsInNewWorkbook = 10

    ' ���ô��ڱ���
    Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"

    ' ���ض���Excel������
    Set wb = Excelapp.Workbooks.Open(App.Path & "\��ӡģ��\����\shbj.xls")
    ' �����һ��������
    Set sh = wb.Sheets(1)
    sh.Activate

    ' ��ѯ���������
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE �䷽���='" & DH & "' ORDER BY ��������,�����"
    DT1.Refresh

    ' ������ݼ�Ϊ�գ����˳�����
    If DT1.Recordset.EOF Then
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
    End If

    ' ��乤��������
    sh.Cells(2, 2) = DT1.Recordset.Fields(0).value
    sh.Cells(2, 4) = DT1.Recordset.Fields(1).value
    sh.Cells(2, 7) = DT1.Recordset.Fields(2).value
    sh.Cells(2, 9) = DT1.Recordset.Fields(3).value
    sh.Cells(3, 12) = DT1.Recordset.Fields(8).value

    ' ��ѯ��������
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE �䷽���='" & DH & "' and Ⱦ������ = '����' ORDER BY ��������,�����"
    DT1.Refresh

    ' �����������
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        i = 5
        Do While Not DT1.Recordset.EOF
            sh.Cells(i, 1) = DT1.Recordset.Fields(6).value
            sh.Cells(i, 3) = DT1.Recordset.Fields(7).value
            sh.Cells(i, 4) = DT1.Recordset.Fields(8).value
            sh.Cells(i, 5) = Format(DT1.Recordset.Fields(15).value, "0.00")
            sh.Cells(i, 6) = Format(DT1.Recordset.Fields(16).value, "0.00")
            DT1.Recordset.MoveNext
            i = i + 1
        Loop
    End If

    ' ��ѯȾ������
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE �䷽���='" & DH & "' and Ⱦ������ = 'Ⱦ��' ORDER BY ��������,�����"
    DT1.Refresh

    ' ���Ⱦ������
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        i = 24
        Do While Not DT1.Recordset.EOF
            sh.Cells(i, 1) = DT1.Recordset.Fields(6).value
            sh.Cells(i, 3) = DT1.Recordset.Fields(7).value
            sh.Cells(i, 4) = DT1.Recordset.Fields(8).value
            sh.Cells(i, 5) = Format(DT1.Recordset.Fields(15).value, "0.00")
            sh.Cells(i, 6) = Format(DT1.Recordset.Fields(16).value, "0.00")
            DT1.Recordset.MoveNext
            i = i + 1
        Loop
    End If

    ' ��ʾԤ�����˳�
    Excelapp.Visible = True
    Excelapp.DisplayAlerts = False
    wb.Sheets.PrintPreview
    ' �رչ����������������
    wb.Close False
    ' �˳�ExcelӦ��
    Excelapp.Quit
    Set Excelapp = Nothing

End Sub




Public Sub pldd4(DT1 As Adodc, dt2 As Adodc, dt3 As Adodc, DH As String) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\pld.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''��¼��
dt3.RecordSource = "SELECT DISTINCT ��������  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT �ϵ����,����,������Ϣ,ѹ��,������� as ��ɫ,�䷽�� as ɫ��,Ⱦ�������� as ��̨,��� as �ͻ�,����,���ϴ�ӡԱ,���ȷ��  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT ��������,ˮ��  FROM pldd WHERE �ϵ����='" & DH & "' group by ��������,ˮ�� Order BY ��������"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "�ӱ����˷������޹�˾���ϵ�"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With
        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "����Ա"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "��̨"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "����/ƥ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "��ɫ"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "ɫ��"

        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + dt3.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '��������
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With


        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "���̿�"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 6
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽����"   '�䷽����
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "�䷽��λ"       '��λ
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "����"       '����
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "��λ"    '�䷽��λ
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With


L = 7
        
Do While Not DT1.Recordset.EOF

If L > 7 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT ԡ��,Ⱦ��������,��������,���ϵ�λ,�䷽,�䷽��λ,У��ֵ,����,����  FROM pldd WHERE �ϵ����='" & DH & "' and ��������='" & DT1.Recordset.Fields(0) & "'  order BY �����"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
If L > 35 Then
i = i + 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(25 * i + 1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "�ӱ����˷������޹�˾���ϵ�"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With

        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "����Ա"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "��̨"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "����/ƥ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "��ɫ"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "ɫ��"

        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.00") + "/" + dt3.Recordset.Fields(10)
        
    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '��������
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With

        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "���̿�"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

L = 6
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "Уֵ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
      '  Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽����"     ''�䷽��λ
      '  Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "�䷽��λ"  '��λ
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "��λ"
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With

L = 7
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)
End If
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#,##0.00")
      '   Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0000")
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '��������
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop
'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '�ر�EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
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


Public Sub pldd44(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, bz As String) ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

   '     On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\pld.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''��¼��
dt3.RecordSource = "SELECT DISTINCT ��������  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT �ϵ����,����,������Ϣ,ѹ��,������� as ��ɫ,�䷽�� as ɫ��,Ⱦ�������� as ��̨,��� as �ͻ�,����,���ϴ�ӡԱ  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT ��������,ˮ��  FROM pldd WHERE �ϵ����='" & DH & "' group by ��������,ˮ�� Order BY ��������"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "�ӱ����˷������޹�˾���ϵ�"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With
        
        Excelapp.ActiveSheet.Cells(50 * i + 2, 2) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 6) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(50 * i + 2, 8) = "����Ա"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 3) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 4) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 5) = "��̨"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 6) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 7) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 8) = "��ɫ"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 9) = "ɫ��"

        Excelapp.ActiveSheet.Cells(50 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0")
        Excelapp.ActiveSheet.Cells(50 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 3, 2), Excelapp.Sheets(1).Cells(50 * i + 4, 9)).Select '��������
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With


        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 3, 2), Excelapp.ActiveSheet.Cells(50 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "���̿�"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 4, 2), Excelapp.Sheets(1).Cells(50 * i + 5, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 6
'        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = "Уֵ"
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "����"       '6
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽��λ"    ''7
        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "Ⱦ��������"  '3
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "�䷽����"    '4
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "��λ"         '5
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With


L = 7
        
Do While Not DT1.Recordset.EOF

If L > 7 Then
Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "����"
Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1) + "L"

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 2), Excelapp.Sheets(1).Cells(50 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(50 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT ԡ��,Ⱦ��������,��������,���ϵ�λ,�䷽,�䷽��λ,У��ֵ,����,����  FROM pldd WHERE �ϵ����='" & DH & "' and ��������='" & DT1.Recordset.Fields(0) & "'  order BY �����"
dt2.Refresh

Excelapp.ActiveSheet.Cells(50 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
If L > 51 Then
i = i + 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 1, 2), Excelapp.ActiveSheet.Cells(50 * i + 1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "�ӱ����˷������޹�˾���ϵ�"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With

        Excelapp.ActiveSheet.Cells(50 * i + 2, 2) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 6) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(50 * i + 2, 8) = "����Ա"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 3) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 4) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 5) = "��̨"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 6) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 7) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 8) = "��ɫ"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 9) = "ɫ��"

        Excelapp.ActiveSheet.Cells(50 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 9) = dt3.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0000")
        
    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 3, 2), Excelapp.Sheets(1).Cells(50 * i + 4, 9)).Select '��������
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With

        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 3, 2), Excelapp.ActiveSheet.Cells(50 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "���̿�"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1) + "L"

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 2), Excelapp.Sheets(1).Cells(50 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

L = 6
'        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = "Уֵ"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 9) = "����"
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "����"       '6
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽��λ"    ''7
        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "Ⱦ��������"  '3
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "�䷽����"    '4
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "��λ"         '5
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With

L = 7
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = dt2.Recordset.Fields(0)
End If
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(50 * i + L, 3) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0000")  '6
         Excelapp.ActiveSheet.Cells(50 * i + L, 4) = dt2.Recordset.Fields(5)                                                '7
         Excelapp.ActiveSheet.Cells(50 * i + L, 5) = Trim(dt2.Recordset.Fields(1))                                         '3
         Excelapp.ActiveSheet.Cells(50 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.0000")                  '4
      '   Excelapp.ActiveSheet.Cells(50 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(50 * i + L, 7) = dt2.Recordset.Fields(3)                                             '5
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = dt2.Recordset.Fields(7)
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 3), Excelapp.Sheets(1).Cells(50 * i + L, 7)).Select '��������
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop



L = L + 2
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
Excelapp.Selection.Merge
Excelapp.Selection.value = "��ע��"
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With

With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = bz
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With

'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '�ر�EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
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

Public Sub pldd(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, bz As String, xs As String, DT4 As Adodc, qx) ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim bhxx As String


        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\pld.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''��¼��
dt3.RecordSource = "SELECT DISTINCT ��������  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT �ϵ����,����,������Ϣ,ѹ��,������� as ��ɫ,�䷽�� as ɫ��,Ⱦ�������� as ��̨,��� as �ͻ�,����,���ϴ�ӡԱ,���ȷ��  FROM pldd WHERE �ϵ����='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT ��������,ˮ��  FROM pldd WHERE �ϵ����='" & DH & "' group by ��������,ˮ�� Order BY ��������"
DT1.Refresh

i = 0
        
        Excelapp.ActiveSheet.Cells(25 * i + 1, 2) = "���˷�֯���޹�˾���ϵ�"
        Excelapp.ActiveSheet.Cells(25 * i + 1, 7) = Trim(DH) + "J"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = Now
        'Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "����Ա"     '''''ԭ���
        'Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 2) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 2) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "��̨"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = "����/ƥ��"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "��ɫ"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "ɫ��"

        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = dt3.Recordset.Fields(2) ''������Ϣ
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = dt3.Recordset.Fields(1) ''����
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = dt3.Recordset.Fields(6) ''��̨
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(7) ''�ͻ�
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + Format(dt3.Recordset.Fields(10), "#0.0") ''����ƥ��
        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(4) ''��ɫ
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = dt3.Recordset.Fields(5) ''ɫ��
        
DT4.RecordSource = "select distinct ���� from kpd where ���� in(select distinct ���� from pld where ���='" & DH & "')"
DT4.Refresh
If Not DT4.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(3, 10) = "��ͬ��"
        Excelapp.ActiveSheet.Cells(4, 10) = DT4.Recordset.Fields(0)
End If
        

L = 9
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 10
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽����"   '�䷽����
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "�䷽��λ"       '��λ
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "����"       '����
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "��λ"    '�䷽��λ
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With

L = 11
        
Do While Not DT1.Recordset.EOF

If L > 11 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT ԡ��,Ⱦ��������,��������,���ϵ�λ,�䷽,�䷽��λ,У��ֵ,����,����,Ⱦ������  FROM pldd WHERE �ϵ����='" & DH & "' and ��������='" & DT1.Recordset.Fields(0) & "'  order BY �����"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         'If InStr(dt2.Recordset.Fields(9), "����") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
        ' Else
        ' Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.name = "����"
        ' Excelapp.ActiveSheet.Range(25 * i + L, 6).Font.Bold = wdToggle
        ' Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.Size = 14
       '  Excelapp.ActiveSheet.Cells(25 * i + L, 6).Value = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
        ' End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         If InStr(dt2.Recordset.Fields(9), "����") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00000")
         End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '��������
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
'L = L + 1
DT1.Recordset.MoveNext
Loop


L = L + 1
'With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
'Excelapp.Selection.Merge
'Excelapp.Selection.Value = "��ע��"
'Excelapp.Selection.Borders.LineStyle = xlContinuous
'End With
If Len(Trim(bz)) > 2 Then
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 3, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = bz
Excelapp.Selection.Font.Size = 8
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If

L = L + 1

DT4.RecordSource = "select * from bgxx where ���ϱ��='" & DH & "'"
DT4.Refresh
If Not DT4.Recordset.EOF Then
DT4.Recordset.MoveFirst
bhxx = ""
Do While Not DT4.Recordset.EOF
bhxx = bhxx + DT4.Recordset.Fields(1) + "/"
DT4.Recordset.MoveNext
Loop
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = "������Ϣ��" + Trim(bhxx)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If



        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

Set Excelapp = Nothing
Excelapp.Quit


End Sub


Public Sub plda(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, DT4 As Adodc) ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim bhxx As String
        Dim gh As String


        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\pld.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM plda WHERE �ϵ����='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''��¼��
dt3.RecordSource = "SELECT DISTINCT ��������  FROM plda WHERE �ϵ����='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT �ϵ����,����,������Ϣ,ѹ��,������� as ��ɫ,�䷽�� as ɫ��,Ⱦ�������� as ��̨,��� as �ͻ�,����,���ϴ�ӡԱ,���ȷ��  FROM plda WHERE �ϵ����='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT ��������,ˮ��  FROM plda WHERE �ϵ����='" & DH & "' group by ��������,ˮ�� Order BY ��������"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
'        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 6)).Select
'        Excelapp.Selection.Merge
'        Excelapp.Selection.Value = "��ׯ���ɼ������޹�˾���ϵ�"
'        Excelapp.Selection.Font.Bold = True
'        Excelapp.Selection.Font.Size = 16
'        End With
        
        Excelapp.ActiveSheet.Cells(25 * i + 1, 2) = "�ӱ����˷������޹�˾���ϵ�"
        Excelapp.ActiveSheet.Cells(25 * i + 1, 7) = "*" + Trim(DH) + "J*"
        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "��Ϣ"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "��ӡ����"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Now
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "����Ա"     '''''ԭ���
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "���ϱ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "Ʒ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "��̨"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "����/ƥ��"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "��ɫ"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "ɫ��"
        gh = dt3.Recordset.Fields(1)                 '''''����'
        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + Format(dt3.Recordset.Fields(10), "#0.0")
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    
DT4.RecordSource = "select distinct ���� from kpd where ���� in(select distinct ���� from pld where ���='" & DH & "')"
DT4.Refresh
If Not DT4.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(3, 10) = "��ͬ��"
        Excelapp.ActiveSheet.Cells(4, 10) = DT4.Recordset.Fields(0)
End If
   
'����������
'    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '��������
'    Excelapp.Selection.Borders.LineStyle = xlContinuous
'    End With


'        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
'        Excelapp.Selection.Merge
'        Excelapp.Selection.Value = "���̿�"
'        Excelapp.Selection.Borders.LineStyle = xlContinuous
'        End With


        Excelapp.ActiveSheet.Cells(6, 2) = "��ѭ��ʱ��"
        Excelapp.ActiveSheet.Cells(6, 4) = "����λ�ã�"
        Excelapp.ActiveSheet.Cells(6, 6) = "��    �ã�"
        Excelapp.ActiveSheet.Cells(7, 2) = "ˮϴ������"
        Excelapp.ActiveSheet.Cells(7, 4) = "�᲼���٣�"
        Excelapp.ActiveSheet.Cells(7, 6) = "�� ˮ �ʣ�"
       ' Excelapp.ActiveSheet.Cells(7, 7) = xs


L = 9
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 10
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽����"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "�䷽��λ"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "ԡ��"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "Ⱦ��������"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "�䷽"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "�䷽����"   '�䷽����
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "Уֵ"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "�䷽��λ"       '��λ
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "����"       '����
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "��λ"    '�䷽��λ
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "����"
        End With


L = 11
        


Do While Not DT1.Recordset.EOF

If L > 11 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "����"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '��������
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT ԡ��,Ⱦ��������,��������,���ϵ�λ,�䷽,�䷽��λ,У��ֵ,����,����,Ⱦ������  FROM plda WHERE �ϵ����='" & DH & "' and ��������='" & DT1.Recordset.Fields(0) & "'  order BY �����"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         If InStr(dt2.Recordset.Fields(9), "����") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.0")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.name = "����"
         Excelapp.ActiveSheet.Range(25 * i + L, 6).Font.Bold = wdToggle
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.Size = 14
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).value = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
         End If
      '   Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         If InStr(dt2.Recordset.Fields(9), "����") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00000")
         End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '��������
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop


L = L + 2
'With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
'Excelapp.Selection.Merge
'Excelapp.Selection.Value = "��ע��"
'Excelapp.Selection.Borders.LineStyle = xlContinuous
'End With
DT4.RecordSource = "select �ܱ�ע from sczy_z  where ���� in(select distinct ���� from kpd where ����='" & gh & "' and len(isnull(����,0))>0)"
DT4.Refresh


If Not DT4.Recordset.EOF Then
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 3, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = DT4.Recordset.Fields(0)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If

L = L + 5

DT4.RecordSource = "select * from bgxx where ���ϱ��='" & DH & "'"
DT4.Refresh
If Not DT4.Recordset.EOF Then
DT4.Recordset.MoveFirst
bhxx = ""
Do While Not DT4.Recordset.EOF
bhxx = bhxx + DT4.Recordset.Fields(1) + "/"
DT4.Recordset.MoveNext
Loop
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = "������Ϣ��" + Trim(bhxx)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If
'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '�ر�EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
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




