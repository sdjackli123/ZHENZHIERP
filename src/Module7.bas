Attribute VB_Name = "Module7"
Public Sub blb(DT1 As Adodc, DH As String)    ''''���ϲ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\blb.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

       ' Excelapp.Selection.Font.FontStyle = "Bold"
DT1.RecordSource = "SELECT ����,��������,���Ϲ��,���ϵ�λ,������ɫ,��������,��������,���Ͽ��� from DHCLB WHERE ����='" & DH & "' order by ���Ͽ���"
DT1.Refresh

Excelapp.ActiveSheet.Cells(4, 2) = DH

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 7
Do While Not DT1.Recordset.EOF
For j = 1 To 7
        Excelapp.ActiveSheet.Cells(i, j) = DT1.Recordset.Fields(j)
Next
DT1.Recordset.MoveNext
i = i + 1
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

