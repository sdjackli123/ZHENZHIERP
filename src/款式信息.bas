Attribute VB_Name = "��ʽ��Ϣ"
Public Sub ksdy(Flex As MSFlexGrid, bt As String)    ''''�ޱ���

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "����������֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\����\cjbb.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k + 1

                          For J = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, J) = "'" & .TextMatrix(i - 1, J)
                      
                          Next J
               
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = bt


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


Public Sub ypdy(Flex As MSFlexGrid, bt As String)    ''''�ޱ���

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "����������֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open ("e:\Excel\����\cjbb.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k + 1

                          For J = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, J) = "'" & .TextMatrix(i - 1, J)
                      
                          Next J
               
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = bt


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


