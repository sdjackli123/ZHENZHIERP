Attribute VB_Name = "染化材料"
Public Sub rhlrk(DT1 As Adodc, DH As String)  ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\rhlrk.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT 供应单位,名称,入库数量,单价,合计金额,入库时间,含税率  FROM mx WHERE 单据号='" & DH & "' order BY IP"
DT1.Refresh
DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(5))
        Excelapp.ActiveSheet.Cells(3, 11) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = "公斤"
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

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub



Public Sub rhlck(DT1 As Adodc, DH As String)  ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\rhlck.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT 出库单位,名称,出库数量,单价,合计金额,出库时间,含税率  FROM ckmx WHERE 单据号='" & DH & "' order BY IP"
DT1.Refresh
DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(5))
        Excelapp.ActiveSheet.Cells(3, 11) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = "公斤"
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

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub


