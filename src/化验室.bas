Attribute VB_Name = "������"
Public Sub bjd(Flex As VSFlexGrid, BT)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\bj.xls")
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


Public Sub gydy(DT1 As Adodc, dt2 As Data, dt3 As Data, pfbh As String)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���տ�.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "select xs from zh where dh='" & pfbh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(11, 14) = DT1.Recordset.Fields(0)
End If

DT1.RecordSource = "select * from pfd where ���='" & pfbh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(pfbh)

        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(2) + "-" + DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(5, 9) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(5, 15) = Trim(DT1.Recordset.Fields(6))
        Excelapp.ActiveSheet.Cells(36, 4) = DT1.Recordset.Fields(5)
        
dt2.RecordSource = "select distinct �������� from pfda where �䷽���='" & pfbh & "' order by ��������"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
i = 11
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(0)
dt3.RecordSource = "select Ⱦ��������,�䷽,У��ֵ,���� from pfda where �䷽���='" & pfbh & "' and ��������='" & dt2.Recordset.Fields(0) & "' order by �����"
dt3.Refresh
If Not dt3.Recordset.EOF Then

dt3.Recordset.MoveFirst
Do While Not dt3.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 2) = dt3.Recordset.Fields(0)
        If Mid(Trim(dt3.Recordset.Fields(1)), 1) = "." Then
        Excelapp.ActiveSheet.Cells(i, 7) = "0" + Trim(dt3.Recordset.Fields(1))
        Else
        Excelapp.ActiveSheet.Cells(i, 7) = Trim(dt3.Recordset.Fields(1))
        End If
        Excelapp.ActiveSheet.Cells(i, 9) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 11) = dt3.Recordset.Fields(2)
        
dt3.Recordset.MoveNext
i = i + 1
Loop
End If
dt2.Recordset.MoveNext
Loop
i = i + 1
End If
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



