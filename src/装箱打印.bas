Attribute VB_Name = "װ���ӡ"
Public Sub xsmxdy(dt1 As Data, dt2 As Data, BH As String) ''''������ϸ��ӡ

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "����������֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\����\XSZXD.xls")
'5)���õ�2��������Ϊ�������

dt1.RecordSource = "SELECT �ͻ�,���,����,FORMAT(SUM(VAL(���)),'#0.00'),FORMAT(SUM(VAL(С��)),'#0') FROM ZXD WHERE ���='" & BH & "' GROUP BY �ͻ�,���,����"
dt1.Refresh

If Not dt1.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(1, 1) = "ǣ���޷������޹�˾�����굥"
        Excelapp.ActiveSheet.Cells(2, 1) = "����"
        Excelapp.ActiveSheet.Cells(2, 7) = "����"
        Excelapp.ActiveSheet.Cells(2, 13) = "���"
        
        Excelapp.ActiveSheet.Cells(2, 2) = dt1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(2, 9) = Trim(dt1.Recordset.Fields(2))
        Excelapp.ActiveSheet.Cells(2, 15) = dt1.Recordset.Fields(1)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
dt2.RecordSource = "SELECT *  FROM ZXD WHERE ���='" & BH & "' order by ���"
dt2.Refresh

i = 1
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(4 + i, 1) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(4 + i, 2) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(4 + i, 3) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(4 + i, 4) = dt2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(4 + i, 5) = dt2.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(4 + i, 6) = dt2.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(4 + i, 7) = dt2.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(4 + i, 8) = dt2.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(4 + i, 9) = dt2.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(4 + i, 10) = dt2.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(4 + i, 11) = dt2.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(4 + i, 12) = dt2.Recordset.Fields(12)
        Excelapp.ActiveSheet.Cells(4 + i, 13) = dt2.Recordset.Fields(13)
        Excelapp.ActiveSheet.Cells(4 + i, 14) = dt2.Recordset.Fields(14)
        Excelapp.ActiveSheet.Cells(4 + i, 15) = dt2.Recordset.Fields(15)
        Excelapp.ActiveSheet.Cells(4 + i, 16) = dt2.Recordset.Fields(18)
        Excelapp.ActiveSheet.Cells(4 + i, 17) = dt2.Recordset.Fields(19)
i = i + 1
dt2.Recordset.MoveNext
Loop
End If

With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 14)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "�ϼ�"
End With

        Excelapp.ActiveSheet.Cells(4 + i, 15) = dt1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(4 + i, 16) = "/"
        Excelapp.ActiveSheet.Cells(4 + i, 17) = dt1.Recordset.Fields(3)
i = i + 1

        Excelapp.ActiveSheet.Cells(4 + i, 1) = "�˷�"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 16)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "�� ��"
End With

        
i = i + 1

        Excelapp.ActiveSheet.Cells(4 + i, 1) = "����д"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = ""
End With
        
i = i + 1

         Excelapp.ActiveSheet.Cells(4 + i, 1) = "��ע"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = ""
End With

i = i + 1
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "�����ջ�֮���������ڣ��뼰ʱ������Ʒ�쳣���֪ͨ��˾���糬�����գ�����Ϊ�����յ���Ʒ���һ�Ʒ�������"
End With

i = i + 1

         
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "��Ʒ������ԭ��Ҫ��������������ߣ�400-6072-876 0536-6235268 ���棺0536-6236109"

End With

i = i + 1

         Excelapp.ActiveSheet.Cells(4 + i, 1) = "�������ڣ��Ƶ��ˣ���"
         Excelapp.ActiveSheet.Cells(4 + i, 8) = "�ֿ�ȷ�ϣ�"
        
         
       
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



Public Sub fhmxdy(dt1 As Data, dt2 As Data, BH As String) ''''������ϸ��ӡ

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "����������֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\����\FHZXD.xls")
'5)���õ�2��������Ϊ�������

dt1.RecordSource = "SELECT * FROM LSFH WHERE ���ݺ�='" & BH & "'"
dt1.Refresh

If Not dt1.Recordset.EOF Then
dt1.RecordSource = "SELECT ������λ,���ݺ�,����,��λ,������,sum(����) FROM LSFH WHERE ���ݺ�='" & BH & "' GROUP BY ������λ,���ݺ�,����,��λ,������ order by ������"
dt1.Refresh

XS = dt1.Recordset.RecordCount

If Not dt1.Recordset.EOF Then
        
dt1.Recordset.MoveFirst

        Excelapp.ActiveSheet.Cells(1, 1) = "ǣ���޷������޹�˾����װ���굥"
        Excelapp.ActiveSheet.Cells(2, 2) = "���" + BH
        Excelapp.ActiveSheet.Cells(2, 6) = "����" + Trim(dt1.Recordset.Fields(2))

i = 0
Do While Not dt1.Recordset.EOF
                                                             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Excelapp.ActiveSheet.Cells(i + 3, 1) = "���" + Trim(dt1.Recordset.Fields(4))
        Excelapp.ActiveSheet.Cells(i + 4, 1) = "���"
        Excelapp.ActiveSheet.Cells(i + 4, 2) = "���"
        Excelapp.ActiveSheet.Cells(i + 4, 3) = "��ɫ"
        Excelapp.ActiveSheet.Cells(i + 4, 4) = "��λ"
        Excelapp.ActiveSheet.Cells(i + 4, 5) = "����"
        

dt2.RecordSource = "SELECT ���,�ͺ�,���,��λ,sum(����)  FROM LSFH WHERE ������='" & dt1.Recordset.Fields(4) & "' group by ���,�ͺ�,���,��λ order by ���,�ͺ�,���"
dt2.Refresh

dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(5 + i, 1) = dt2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(5 + i, 2) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(5 + i, 3) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(5 + i, 4) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(5 + i, 5) = dt2.Recordset.Fields(4)
        
i = i + 1
dt2.Recordset.MoveNext
Loop
i = i + 2
dt1.Recordset.MoveNext
Loop
End If

i = i + 3

dt1.RecordSource = "SELECT sum(����) FROM LSFH WHERE ���ݺ�='" & BH & "'"
dt1.Refresh

        Excelapp.ActiveSheet.Cells(i + 2, 1) = "�ϼ�������" + Trim(XS)
        Excelapp.ActiveSheet.Cells(i + 2, 6) = "�ϼƼ�����" + Trim(dt1.Recordset.Fields(0))
        
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







