Attribute VB_Name = "Module7"
Public Sub blb(DT1 As Adodc, DH As String)    ''''备料操作

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\blb.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

       ' Excelapp.Selection.Font.FontStyle = "Bold"
DT1.RecordSource = "SELECT 单号,材料名称,材料规格,材料单位,材料颜色,材料批号,材料数量,材料库类 from DHCLB WHERE 单号='" & DH & "' order by 材料库类"
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

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

