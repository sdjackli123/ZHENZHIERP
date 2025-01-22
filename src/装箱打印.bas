Attribute VB_Name = "装箱打印"
Public Sub xsmxdy(dt1 As Data, dt2 As Data, BH As String) ''''销售明细打印

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军制衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("e:\Excel\成衣\XSZXD.xls")
'5)设置第2个工作表为活动工作表：

dt1.RecordSource = "SELECT 客户,编号,日期,FORMAT(SUM(VAL(金额)),'#0.00'),FORMAT(SUM(VAL(小计)),'#0') FROM ZXD WHERE 编号='" & BH & "' GROUP BY 客户,编号,日期"
dt1.Refresh

If Not dt1.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(1, 1) = "牵手娃服饰有限公司发货详单"
        Excelapp.ActiveSheet.Cells(2, 1) = "家人"
        Excelapp.ActiveSheet.Cells(2, 7) = "日期"
        Excelapp.ActiveSheet.Cells(2, 13) = "编号"
        
        Excelapp.ActiveSheet.Cells(2, 2) = dt1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(2, 9) = Trim(dt1.Recordset.Fields(2))
        Excelapp.ActiveSheet.Cells(2, 15) = dt1.Recordset.Fields(1)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
dt2.RecordSource = "SELECT *  FROM ZXD WHERE 编号='" & BH & "' order by 款号"
dt2.Refresh

i = 1
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(4 + i, 1) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(4 + i, 2) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(4 + i, 3) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(4 + i, 4) = dt2.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(4 + i, 5) = dt2.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(4 + i, 6) = dt2.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(4 + i, 7) = dt2.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(4 + i, 8) = dt2.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(4 + i, 9) = dt2.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(4 + i, 10) = dt2.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(4 + i, 11) = dt2.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(4 + i, 12) = dt2.Recordset.Fields(12)
        Excelapp.ActiveSheet.Cells(4 + i, 13) = dt2.Recordset.Fields(13)
        Excelapp.ActiveSheet.Cells(4 + i, 14) = dt2.Recordset.Fields(14)
        Excelapp.ActiveSheet.Cells(4 + i, 15) = dt2.Recordset.Fields(15)
        Excelapp.ActiveSheet.Cells(4 + i, 16) = dt2.Recordset.Fields(18)
        Excelapp.ActiveSheet.Cells(4 + i, 17) = dt2.Recordset.Fields(19)
i = i + 1
dt2.Recordset.MoveNext
Loop
End If

With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 14)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "合计"
End With

        Excelapp.ActiveSheet.Cells(4 + i, 15) = dt1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(4 + i, 16) = "/"
        Excelapp.ActiveSheet.Cells(4 + i, 17) = dt1.Recordset.Fields(3)
i = i + 1

        Excelapp.ActiveSheet.Cells(4 + i, 1) = "运费"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 16)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "共 箱"
End With

        
i = i + 1

        Excelapp.ActiveSheet.Cells(4 + i, 1) = "金额大写"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = ""
End With
        
i = i + 1

         Excelapp.ActiveSheet.Cells(4 + i, 1) = "备注"
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 2), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = ""
End With

i = i + 1
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "自您收货之日起两日内，请及时将到货品异常情况通知我司。如超出两日，则视为您已收到货品，且货品完好无损，"
End With

i = i + 1

         
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(4 + i, 1), Excelapp.ActiveSheet.Cells(4 + i, 17)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.Value = "货品数量与原定要求相符。服务热线：400-6072-876 0536-6235268 传真：0536-6236109"

End With

i = i + 1

         Excelapp.ActiveSheet.Cells(4 + i, 1) = "销售内勤（制单人）："
         Excelapp.ActiveSheet.Cells(4 + i, 8) = "仓库确认："
        
         
       
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



Public Sub fhmxdy(dt1 As Data, dt2 As Data, BH As String) ''''发货明细打印

        Dim i   As Integer
        Dim J   As Integer
        Dim k   As Integer
        Dim X   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "凤军制衣软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open ("e:\Excel\成衣\FHZXD.xls")
'5)设置第2个工作表为活动工作表：

dt1.RecordSource = "SELECT * FROM LSFH WHERE 单据号='" & BH & "'"
dt1.Refresh

If Not dt1.Recordset.EOF Then
dt1.RecordSource = "SELECT 购货单位,单据号,日期,单位,发货地,sum(数量) FROM LSFH WHERE 单据号='" & BH & "' GROUP BY 购货单位,单据号,日期,单位,发货地 order by 发货地"
dt1.Refresh

XS = dt1.Recordset.RecordCount

If Not dt1.Recordset.EOF Then
        
dt1.Recordset.MoveFirst

        Excelapp.ActiveSheet.Cells(1, 1) = "牵手娃服饰有限公司发货装箱详单"
        Excelapp.ActiveSheet.Cells(2, 2) = "编号" + BH
        Excelapp.ActiveSheet.Cells(2, 6) = "日期" + Trim(dt1.Recordset.Fields(2))

i = 0
Do While Not dt1.Recordset.EOF
                                                             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Excelapp.ActiveSheet.Cells(i + 3, 1) = "箱号" + Trim(dt1.Recordset.Fields(4))
        Excelapp.ActiveSheet.Cells(i + 4, 1) = "款号"
        Excelapp.ActiveSheet.Cells(i + 4, 2) = "规格"
        Excelapp.ActiveSheet.Cells(i + 4, 3) = "颜色"
        Excelapp.ActiveSheet.Cells(i + 4, 4) = "单位"
        Excelapp.ActiveSheet.Cells(i + 4, 5) = "数量"
        

dt2.RecordSource = "SELECT 款号,型号,规格,单位,sum(数量)  FROM LSFH WHERE 发货地='" & dt1.Recordset.Fields(4) & "' group by 款号,型号,规格,单位 order by 款号,型号,规格"
dt2.Refresh

dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(5 + i, 1) = dt2.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(5 + i, 2) = dt2.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(5 + i, 3) = dt2.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(5 + i, 4) = dt2.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(5 + i, 5) = dt2.Recordset.Fields(4)
        
i = i + 1
dt2.Recordset.MoveNext
Loop
i = i + 2
dt1.Recordset.MoveNext
Loop
End If

i = i + 3

dt1.RecordSource = "SELECT sum(数量) FROM LSFH WHERE 单据号='" & BH & "'"
dt1.Refresh

        Excelapp.ActiveSheet.Cells(i + 2, 1) = "合计箱数：" + Trim(XS)
        Excelapp.ActiveSheet.Cells(i + 2, 6) = "合计件数：" + Trim(dt1.Recordset.Fields(0))
        
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







