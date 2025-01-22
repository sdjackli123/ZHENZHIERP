Attribute VB_Name = "Module3"
Public Sub DGYDOutadodcToExcel(DT1 As Adodc, dt2 As Adodc, DH As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\DGYD.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

dt2.RecordSource = "SELECT SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_ZDH.日期,SCZY_ZDH.负责人,sum(cmb.数量),sum(val(cmb.计划))  FROM SCZY_ZDH,cmb WHERE SCZY_ZDH.单号='" & DH & "' and cmb.单号=SCZY_ZDH.单号 GROUP BY SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_ZDH.款式,SCZY_ZDH.日期,SCZY_ZDH.负责人"
dt2.Refresh
        Excelapp.ActiveSheet.Cells(5, 2) = dt2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(5, 5) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(6, 2) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(5, 11) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(6, 11) = dt2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(5, 8) = dt2.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(6, 8) = dt2.Recordset.Fields(6)

DT1.RecordSource = "SELECT distinct 款号  FROM cmb WHERE 单号='" & DH & "'"
DT1.Refresh
L = DT1.Recordset.RecordCount

If L < 1 Then Exit Sub
i = 8
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT 款号,sum(数量),sum(val(计划)) FROM cmb WHERE 单号='" & DH & "' AND 款号='" & DT1.Recordset.Fields(0) & "' GROUP BY 款号"
dt2.Refresh
i = i + 1
        Excelapp.ActiveSheet.Cells(i, 1) = "款号"
        Excelapp.ActiveSheet.Cells(i, 2) = "订单数量"
        Excelapp.ActiveSheet.Cells(i, 3) = "计划数量"
i = i + 1
        Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(i, 2) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = dt2.Recordset.Fields(2)

i = i + 1
        Excelapp.ActiveSheet.Cells(i, 1) = "颜色"
        Excelapp.ActiveSheet.Cells(i, 2) = "尺码"
        Excelapp.ActiveSheet.Cells(i, 3) = "订单数量"
        Excelapp.ActiveSheet.Cells(i, 4) = "计划数量"

dt2.RecordSource = "select * from cmb WHERE 单号='" & DH & "' AND 款号='" & DT1.Recordset.Fields(0) & "' order by 颜色,尺码"
dt2.Refresh
i = i + 1
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 2) = dt2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 3) = Format(dt2.Recordset.Fields(5), "#0")
        Excelapp.ActiveSheet.Cells(i, 4) = Format(dt2.Recordset.Fields(6), "#0")
i = i + 1
dt2.Recordset.MoveNext
Loop
DT1.Recordset.MoveNext
Loop


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
Public Sub YGYDOutadodcToExcel(DT1 As Adodc, dt2 As Adodc, DH As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\DGYD.xls")
'5)设置第2个工作表为活动工作表：

DT1.RecordSource = "SELECT 款号  FROM SCZY_X WHERE 单号='" & DH & "' GROUP BY 款号"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub
DT1.RecordSource = "SELECT 款号 FROM SCZY_X WHERE 单号='" & DH & "' GROUP BY 款号"
DT1.Refresh
DT1.Recordset.MoveFirst
PP = 1
op = 0
IO = 1
Do While Not DT1.Recordset.EOF

If op > 0 Then
If Int(op / 6) = op / 6 Then
PP = PP + 1 '''''''''''PP表
op = 0   ''''''''''''页
End If
End If
        Excelapp.Sheets(PP).Activate


dt2.RecordSource = "SELECT SCZY_Z.客户,SCZY_Z.单号,SCZY_Z.款式,SCZY_Z.面料,SCZY_Z.数量,SCZY_Z.日期,SCZY_Z.交期,SCZY_Z.负责人  FROM SCZY_Z WHERE 单号='" & DH & "'"
dt2.Refresh
        Excelapp.ActiveSheet.Cells(op * 43 + 4, 8) = "共" + Trim(L) + "页"
        Excelapp.ActiveSheet.Cells(op * 43 + 4, 10) = "第" + Trim(IO) + "页"
        Excelapp.ActiveSheet.Cells(op * 43 + 5, 2) = dt2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(op * 43 + 6, 2) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(op * 43 + 5, 5) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(op * 43 + 6, 5) = dt2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(op * 43 + 5, 8) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(op * 43 + 6, 8) = Trim(dt2.Recordset.Fields(6))
        Excelapp.ActiveSheet.Cells(op * 43 + 5, 11) = Trim(dt2.Recordset.Fields(5))
        Excelapp.ActiveSheet.Cells(op * 43 + 6, 11) = dt2.Recordset.Fields(7)
dt2.RecordSource = "select * from SCZY_X WHERE 单号='" & DH & "' AND 款号='" & DT1.Recordset.Fields(0) & "' order by 序号"
dt2.Refresh

        Excelapp.ActiveSheet.Cells(op * 43 + 16, 1) = dt2.Recordset.Fields(30)
        Excelapp.ActiveSheet.Cells(op * 43 + 16, 7) = dt2.Recordset.Fields(31)
        Excelapp.ActiveSheet.Cells(op * 43 + 23, 3) = dt2.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(op * 43 + 24, 3) = dt2.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(op * 43 + 26, 3) = dt2.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(op * 43 + 29, 3) = dt2.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(op * 43 + 32, 3) = dt2.Recordset.Fields(9)
i = 8
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(op * 43 + i, 1) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 2) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 3) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 4) = dt2.Recordset.Fields(10) + "/" + dt2.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 5) = dt2.Recordset.Fields(12) + "/" + dt2.Recordset.Fields(13)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 6) = dt2.Recordset.Fields(14) + "/" + dt2.Recordset.Fields(15)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 7) = dt2.Recordset.Fields(16) + "/" + dt2.Recordset.Fields(17)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 8) = dt2.Recordset.Fields(18) + "/" + dt2.Recordset.Fields(19)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 9) = dt2.Recordset.Fields(20) + "/" + dt2.Recordset.Fields(21)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 10) = dt2.Recordset.Fields(22) + "/" + dt2.Recordset.Fields(23)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 11) = dt2.Recordset.Fields(24) + "/" + dt2.Recordset.Fields(25)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 12) = dt2.Recordset.Fields(26) + "/" + dt2.Recordset.Fields(27)
        Excelapp.ActiveSheet.Cells(op * 43 + i, 13) = dt2.Recordset.Fields(28) + "/" + dt2.Recordset.Fields(29)
i = i + 1
dt2.Recordset.MoveNext
Loop

op = op + 1
IO = IO + 1
DT1.Recordset.MoveNext
Loop

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


