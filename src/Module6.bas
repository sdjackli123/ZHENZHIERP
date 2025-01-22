Attribute VB_Name = "Module6"

Public Sub BTDY(DT1 As Adodc, DH As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\BTDY.xls")
'5)设置第2个工作表为活动工作表：

DT1.RecordSource = "SELECT distinct SCZY_ZDH.客户,SCZY_ZDH.单号,SCZY_XDH.款号,SCZY_XDH.款式,SCZY_ZDH.日期  FROM SCZY_ZDH,SCZY_XDH WHERE SCZY_XDH.单号=SCZY_ZDH.单号 AND SCZY_ZDH.单号='" & DH & "' ORDER BY 款号"
DT1.Refresh
i = 0
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(2, 2) = "客户"
        Excelapp.ActiveSheet.Cells(2, 3) = "单号"
        Excelapp.ActiveSheet.Cells(2, 4) = "款号"
        Excelapp.ActiveSheet.Cells(2, 5) = "款式"
        Excelapp.ActiveSheet.Cells(2, 6) = "日期"
        
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(3 + i, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3 + i, 3) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(3 + i, 4) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(3 + i, 5) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(3 + i, 6) = Trim(DT1.Recordset.Fields(4))
        
i = i + 1
DT1.Recordset.MoveNext
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




