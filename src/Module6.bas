Attribute VB_Name = "Module6"

Public Sub BTDY(DT1 As Adodc, DH As String) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\BTDY.xls")
'5)���õ�2��������Ϊ�������

DT1.RecordSource = "SELECT distinct SCZY_ZDH.�ͻ�,SCZY_ZDH.����,SCZY_XDH.���,SCZY_XDH.��ʽ,SCZY_ZDH.����  FROM SCZY_ZDH,SCZY_XDH WHERE SCZY_XDH.����=SCZY_ZDH.���� AND SCZY_ZDH.����='" & DH & "' ORDER BY ���"
DT1.Refresh
i = 0
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(2, 2) = "�ͻ�"
        Excelapp.ActiveSheet.Cells(2, 3) = "����"
        Excelapp.ActiveSheet.Cells(2, 4) = "���"
        Excelapp.ActiveSheet.Cells(2, 5) = "��ʽ"
        Excelapp.ActiveSheet.Cells(2, 6) = "����"
        
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(3 + i, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3 + i, 3) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(3 + i, 4) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(3 + i, 5) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(3 + i, 6) = Trim(DT1.Recordset.Fields(4))
        
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




