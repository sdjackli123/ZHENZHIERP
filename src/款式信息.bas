Attribute VB_Name = "款式信息"
Public Sub ksdy(Flex As MSFlexGrid, bt As String)    ''''无标题

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "凤军制衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("e:\Excel\成衣\cjbb.xls")
'5)设置第2个工作表为活动工作表：
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

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


Public Sub ypdy(Flex As MSFlexGrid, bt As String)    ''''无标题

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "凤军制衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("e:\Excel\成衣\cjbb.xls")
'5)设置第2个工作表为活动工作表：
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

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


