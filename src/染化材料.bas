Attribute VB_Name = "Ⱦ������"
Public Sub rhlrk(DT1 As Adodc, DH As String)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\rhlrk.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT ��Ӧ��λ,����,�������,����,�ϼƽ��,���ʱ��,��˰��  FROM mx WHERE ���ݺ�='" & DH & "' order BY IP"
DT1.Refresh
DT1.Recordset.MoveFirst
'DT3.RecordSource = "select ���� from gys where ���='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(5))
        Excelapp.ActiveSheet.Cells(3, 11) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "ҵ��Ա��" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = "����"
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(6)
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



Public Sub rhlck(DT1 As Adodc, DH As String)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\rhlck.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT ���ⵥλ,����,��������,����,�ϼƽ��,����ʱ��,��˰��  FROM ckmx WHERE ���ݺ�='" & DH & "' order BY IP"
DT1.Refresh
DT1.Recordset.MoveFirst
'DT3.RecordSource = "select ���� from gys where ���='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(5))
        Excelapp.ActiveSheet.Cells(3, 11) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "ҵ��Ա��" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = "����"
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(6)
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


