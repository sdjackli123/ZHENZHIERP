Attribute VB_Name = "包装尺码"
Public Sub bzcm(dt1 As Data, kh As String) ''''无标题

        Dim i   As Integer
        Dim J   As Integer
        Dim bkh   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军制衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("e:\Excel\成衣\bztm.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

dt1.RecordSource = "SELECT  * FROM bztm WHERE 卡号='" & kh & "'"
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
          Excelapp.Selection.Font.name = "宋体"
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

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing


End Sub

