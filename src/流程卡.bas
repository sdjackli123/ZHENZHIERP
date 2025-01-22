Attribute VB_Name = "流程卡"
Public Sub cpk(DT1 As Data, DH As String)  ''''无标题

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军成衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("E:\Excel\成衣\cpk.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT *  FROM cpk WHERE 卡号='" & DH & "' order BY 编号"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(2, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(4, 1) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(6, 1) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(8, 1) = Trim(DH)
        Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(4)
        
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 2
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)
        
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(i, 8) = "*" + DT1.Recordset.Fields(8) + "J" + "*"
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

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub


Public Sub cpk1(DT1 As Data, DH As String)  ''''无标题

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军成衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("E:\Excel\成衣\cpk1.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT *  FROM cpk WHERE 卡号='" & DH & "' order BY 编号"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(2, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(4, 1) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(6, 1) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(8, 1) = Trim(DH)
        Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(4)
        
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 2
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)
        
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(i, 8) = "*" + DT1.Recordset.Fields(8) + "J" + "*"
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

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub



Public Sub OutDataToExcel3(Flex As MSFlexGrid, FD1, FD2, FD3, bt) ''''按一字段合计（含标题）

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
Excelapp.Workbooks.Open ("e:\Excel\成衣\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
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
               
         If i >= 3 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, FD1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, FD2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, FD3)) + Q3
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = bt

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, FD1) = Q1
Excelapp.ActiveSheet.Cells(i, FD2) = Q2
Excelapp.ActiveSheet.Cells(i, FD3) = Q3
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


