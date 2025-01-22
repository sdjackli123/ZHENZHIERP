Attribute VB_Name = "财务信息"
Public Sub gxxs(Flex As VSFlexGrid, DT1 As Adodc, DH As String)    ''''系数操作

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\gxxs.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"
DT1.RecordSource = "SELECT distinct cmb.款号  FROM cmb WHERE cmb.款号='" & DH & "'"
DT1.Refresh
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(0)

        With Flex

                k = .Rows


          For i = 6 To k + 6 - 1

                          For j = 1 To .Cols

                              
                              DoEvents
                              
                              Excelapp.ActiveSheet.Cells(i + 2, j) = "'" & .TextMatrix(i - 5, j)
                              
                              If j = 5 Then
                              Excelapp.ActiveSheet.Cells(i + 2, j + 1) = "'" & .TextMatrix(i - 5, j)
                              j = j + 1
                              End If
                              If j = 6 Then
                              j = j + 1
                              End If
                              If j = 7 Then
                              Excelapp.ActiveSheet.Cells(i + 2, j - 2) = "'" & .TextMatrix(i - 5, j)
                              End If
                              
                                      
                          Next j
               
         Next i
        End With

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


Public Sub OutadodcToExcel6(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, BT) ''''按一字段合计（含标题）

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
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
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         Q6 = Val(Excelapp.ActiveSheet.Cells(i, fd6)) + Q6
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5
Excelapp.ActiveSheet.Cells(i, fd6) = Q6

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

Public Sub OutadodcToExcel5(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, BT) ''''按一字段合计（含标题）

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
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
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5

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


Public Sub OutadodcToExcel9(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, fd9, BT) ''''按一字段合计（含标题）

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
        Q7 = 0
        Q8 = 0
        Q9 = 0
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
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         Q6 = Val(Excelapp.ActiveSheet.Cells(i, fd6)) + Q6
         Q7 = Val(Excelapp.ActiveSheet.Cells(i, fd7)) + Q7
         Q8 = Val(Excelapp.ActiveSheet.Cells(i, fd8)) + Q8
         Q9 = Val(Excelapp.ActiveSheet.Cells(i, fd9)) + Q9
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5
Excelapp.ActiveSheet.Cells(i, fd6) = Q6
Excelapp.ActiveSheet.Cells(i, fd7) = Q7
Excelapp.ActiveSheet.Cells(i, fd8) = Q8
Excelapp.ActiveSheet.Cells(i, fd9) = Q9

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

