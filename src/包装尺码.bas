Attribute VB_Name = "��װ����"
Public Sub bzcm(dt1 As Data, kh As String) ''''�ޱ���

        Dim i   As Integer
        Dim J   As Integer
        Dim bkh   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "����������֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\����\bztm.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

dt1.RecordSource = "SELECT  * FROM bztm WHERE ����='" & kh & "'"
dt1.Refresh

If Not dt1.Recordset.EOF Then
        l = 1
        For i = 0 To 49
        If Len(dt1.Recordset.Fields(7 + i)) > 8 And dt1.Recordset.Fields(7 + i) <> "" Then
'Excelapp.ActiveSheet.Rows(l).RowHeight = 0.1 / 0.035
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(l, 1), Excelapp.ActiveSheet.Cells(l, 1)).Select
          Excelapp.Selection.Font.name = "ExtCode39XS"
          Excelapp.Selection.Merge
          Excelapp.Selection.Font.Size = 9
          Excelapp.Selection.Value = "*" + dt1.Recordset.Fields(i + 7) + "J*"
End With
        
        l = l + 1

With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(l, 1), Excelapp.ActiveSheet.Cells(l, 1)).Select
          Excelapp.Selection.Font.name = "����"
          Excelapp.Selection.Merge
          Excelapp.Selection.Font.Size = 9
          Excelapp.Selection.Value = dt1.Recordset.Fields(3)
End With

        l = l + 1
        
        End If
        Next
        
        
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

