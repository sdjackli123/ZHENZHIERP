Attribute VB_Name = "对版室"
Public Sub gd(bg As VSFlexGrid)
  Dim Row, col As Integer                        '定义两个变量用于接收表格的行与列
  With bg
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = "x"       '将表格中的每一个单元格赋值为Row+Col
    Next Row
  End With
End Sub
Public Sub shbj(DT1 As Adodc, dt2 As Adodc, DH As String)
    ' 声明变量
    Dim i As Integer
    Dim Excelapp As Excel.Application
    Dim wb As Excel.Workbook
    Dim sh As Excel.Worksheet

    ' 创建Excel应用实例
    Set Excelapp = New Excel.Application
    ' 新工作簿中包含10个工作表
    Excelapp.SheetsInNewWorkbook = 10

    ' 设置窗口标题
    Excelapp.Caption = "广兴打印模版软件之打印"

    ' 打开特定的Excel工作簿
    Set wb = Excelapp.Workbooks.Open(App.Path & "\打印模版\广兴\shbj.xls")
    ' 激活第一个工作表
    Set sh = wb.Sheets(1)
    sh.Activate

    ' 查询并填充数据
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE 配方编号='" & DH & "' ORDER BY 工序名称,次序号"
    DT1.Refresh

    ' 如果数据集为空，则退出程序
    If DT1.Recordset.EOF Then
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
    End If

    ' 填充工作簿数据
    sh.Cells(2, 2) = DT1.Recordset.Fields(0).value
    sh.Cells(2, 4) = DT1.Recordset.Fields(1).value
    sh.Cells(2, 7) = DT1.Recordset.Fields(2).value
    sh.Cells(2, 9) = DT1.Recordset.Fields(3).value
    sh.Cells(3, 12) = DT1.Recordset.Fields(8).value

    ' 查询助剂数据
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE 配方编号='" & DH & "' and 染化助库 = '助剂' ORDER BY 工序名称,次序号"
    DT1.Refresh

    ' 填充助剂数据
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        i = 5
        Do While Not DT1.Recordset.EOF
            sh.Cells(i, 1) = DT1.Recordset.Fields(6).value
            sh.Cells(i, 3) = DT1.Recordset.Fields(7).value
            sh.Cells(i, 4) = DT1.Recordset.Fields(8).value
            sh.Cells(i, 5) = Format(DT1.Recordset.Fields(15).value, "0.00")
            sh.Cells(i, 6) = Format(DT1.Recordset.Fields(16).value, "0.00")
            DT1.Recordset.MoveNext
            i = i + 1
        Loop
    End If

    ' 查询染料数据
    DT1.RecordSource = "SELECT * FROM v_dpfd_sh_bj WHERE 配方编号='" & DH & "' and 染化助库 = '染料' ORDER BY 工序名称,次序号"
    DT1.Refresh

    ' 填充染料数据
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        i = 24
        Do While Not DT1.Recordset.EOF
            sh.Cells(i, 1) = DT1.Recordset.Fields(6).value
            sh.Cells(i, 3) = DT1.Recordset.Fields(7).value
            sh.Cells(i, 4) = DT1.Recordset.Fields(8).value
            sh.Cells(i, 5) = Format(DT1.Recordset.Fields(15).value, "0.00")
            sh.Cells(i, 6) = Format(DT1.Recordset.Fields(16).value, "0.00")
            DT1.Recordset.MoveNext
            i = i + 1
        Loop
    End If

    ' 显示预览并退出
    Excelapp.Visible = True
    Excelapp.DisplayAlerts = False
    wb.Sheets.PrintPreview
    ' 关闭工作簿，不保存更改
    wb.Close False
    ' 退出Excel应用
    Excelapp.Quit
    Set Excelapp = Nothing

End Sub




Public Sub pldd4(DT1 As Adodc, dt2 As Adodc, dt3 As Adodc, DH As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\pld.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''记录数
dt3.RecordSource = "SELECT DISTINCT 工序名称  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT 料单编号,锅号,生产信息,压力,生产类别 as 颜色,配方单 as 色号,染化助单价 as 车台,审核 as 客户,重量,配料打印员,审核确认  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT 工序名称,水量  FROM pldd WHERE 料单编号='" & DH & "' group by 工序名称,水量 Order BY 工序名称"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "河北广兴服饰有限公司配料单"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With
        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "信息"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "打印日期"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "操作员"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "锅号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "品名"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "车台"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "客户"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "重量/匹数"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "颜色"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "色号"

        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + dt3.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '加网格线
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With


        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "流程卡"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 6
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方用量"   '配方用量
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "配方单位"       '单位
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "比例"       '比例
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "单位"    '配方单位
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With


L = 7
        
Do While Not DT1.Recordset.EOF

If L > 7 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT 浴比,染化助名称,配料用量,配料单位,配方,配方单位,校正值,车速,批次  FROM pldd WHERE 料单编号='" & DH & "' and 工序名称='" & DT1.Recordset.Fields(0) & "'  order BY 次序号"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
If L > 35 Then
i = i + 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(25 * i + 1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "河北广兴服饰有限公司配料单"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With

        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "信息"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "打印日期"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "操作员"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "锅号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "品名"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "车台"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "客户"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "重量/匹数"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "颜色"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "色号"

        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.00") + "/" + dt3.Recordset.Fields(10)
        
    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '加网格线
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With

        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "流程卡"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

L = 6
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "校值"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = "工艺"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
      '  Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方用量"     ''配方单位
      '  Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "配方单位"  '单位
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "比例"
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "单位"
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With

L = 7
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)
End If
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#,##0.00")
      '   Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0000")
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '加网格线
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop
'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '关闭EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
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


Public Sub pldd44(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, bz As String) ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

   '     On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\pld.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''记录数
dt3.RecordSource = "SELECT DISTINCT 工序名称  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT 料单编号,锅号,生产信息,压力,生产类别 as 颜色,配方单 as 色号,染化助单价 as 车台,审核 as 客户,重量,配料打印员  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT 工序名称,水量  FROM pldd WHERE 料单编号='" & DH & "' group by 工序名称,水量 Order BY 工序名称"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "河北广兴服饰有限公司配料单"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With
        
        Excelapp.ActiveSheet.Cells(50 * i + 2, 2) = "信息"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 6) = "打印日期"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(50 * i + 2, 8) = "操作员"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 3) = "锅号"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 4) = "品名"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 5) = "车台"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 6) = "客户"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 7) = "重量"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 8) = "颜色"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 9) = "色号"

        Excelapp.ActiveSheet.Cells(50 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0")
        Excelapp.ActiveSheet.Cells(50 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 3, 2), Excelapp.Sheets(1).Cells(50 * i + 4, 9)).Select '加网格线
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With


        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 3, 2), Excelapp.ActiveSheet.Cells(50 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "流程卡"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 4, 2), Excelapp.Sheets(1).Cells(50 * i + 5, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 6
'        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = "校值"
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "比例"       '6
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方单位"    ''7
        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "染化助名称"  '3
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "配方用量"    '4
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "单位"         '5
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With


L = 7
        
Do While Not DT1.Recordset.EOF

If L > 7 Then
Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "工序"
Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1) + "L"

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 2), Excelapp.Sheets(1).Cells(50 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(50 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT 浴比,染化助名称,配料用量,配料单位,配方,配方单位,校正值,车速,批次  FROM pldd WHERE 料单编号='" & DH & "' and 工序名称='" & DT1.Recordset.Fields(0) & "'  order BY 次序号"
dt2.Refresh

Excelapp.ActiveSheet.Cells(50 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
If L > 51 Then
i = i + 1
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 1, 2), Excelapp.ActiveSheet.Cells(50 * i + 1, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "河北广兴服饰有限公司配料单"
        Excelapp.Selection.Font.Bold = True
        Excelapp.Selection.Font.Size = 16
        End With

        Excelapp.ActiveSheet.Cells(50 * i + 2, 2) = "信息"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 6) = "打印日期"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 7) = Date
        Excelapp.ActiveSheet.Cells(50 * i + 2, 8) = "操作员"
        Excelapp.ActiveSheet.Cells(50 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 3) = "锅号"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 4) = "品名"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 5) = "车台"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 6) = "客户"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 7) = "重量"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 8) = "颜色"
        Excelapp.ActiveSheet.Cells(50 * i + 3, 9) = "色号"

        Excelapp.ActiveSheet.Cells(50 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(50 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 9) = dt3.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(50 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0000")
        
    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + 3, 2), Excelapp.Sheets(1).Cells(50 * i + 4, 9)).Select '加网格线
    Excelapp.Selection.Borders.LineStyle = xlContinuous
    End With

        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + 3, 2), Excelapp.ActiveSheet.Cells(50 * i + 4, 2)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "流程卡"
        Excelapp.Selection.Borders.LineStyle = xlContinuous
        End With

L = 5
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = DT1.Recordset.Fields(1) + "L"

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 2), Excelapp.Sheets(1).Cells(50 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

L = 6
'        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = "校值"
'        Excelapp.ActiveSheet.Cells(50 * i + L, 9) = "工艺"
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = "比例"       '6
        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方单位"    ''7
        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "染化助名称"  '3
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = "配方用量"    '4
      '  Excelapp.ActiveSheet.Cells(50 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = "单位"         '5
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With

L = 7
        Excelapp.ActiveSheet.Cells(50 * i + L, 2) = dt2.Recordset.Fields(0)
End If
'        Excelapp.ActiveSheet.Cells(50 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(50 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(50 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(50 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(50 * i + L, 3) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0000")  '6
         Excelapp.ActiveSheet.Cells(50 * i + L, 4) = dt2.Recordset.Fields(5)                                                '7
         Excelapp.ActiveSheet.Cells(50 * i + L, 5) = Trim(dt2.Recordset.Fields(1))                                         '3
         Excelapp.ActiveSheet.Cells(50 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.0000")                  '4
      '   Excelapp.ActiveSheet.Cells(50 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(50 * i + L, 7) = dt2.Recordset.Fields(3)                                             '5
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(50 * i + L, 8), Excelapp.ActiveSheet.Cells(50 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = dt2.Recordset.Fields(7)
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(50 * i + L, 3), Excelapp.Sheets(1).Cells(50 * i + L, 7)).Select '加网格线
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop



L = L + 2
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
Excelapp.Selection.Merge
Excelapp.Selection.value = "备注："
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With

With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = bz
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With

'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '关闭EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
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

Public Sub pldd(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, bz As String, xs As String, DT4 As Adodc, qx) ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim bhxx As String


        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\pld.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''记录数
dt3.RecordSource = "SELECT DISTINCT 工序名称  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT 料单编号,锅号,生产信息,压力,生产类别 as 颜色,配方单 as 色号,染化助单价 as 车台,审核 as 客户,重量,配料打印员,审核确认  FROM pldd WHERE 料单编号='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT 工序名称,水量  FROM pldd WHERE 料单编号='" & DH & "' group by 工序名称,水量 Order BY 工序名称"
DT1.Refresh

i = 0
        
        Excelapp.ActiveSheet.Cells(25 * i + 1, 2) = "广兴纺织有限公司配料单"
        Excelapp.ActiveSheet.Cells(25 * i + 1, 7) = Trim(DH) + "J"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = "信息"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = "打印日期"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = Now
        'Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "操作员"     '''''原班次
        'Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 2) = "锅号"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 2) = "品名"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "车台"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "客户"
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = "重量/匹数"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "颜色"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "色号"

        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = dt3.Recordset.Fields(2) ''生产信息
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = dt3.Recordset.Fields(1) ''锅号
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = dt3.Recordset.Fields(6) ''车台
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(7) ''客户
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + Format(dt3.Recordset.Fields(10), "#0.0") ''重量匹数
        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(4) ''颜色
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = dt3.Recordset.Fields(5) ''色号
        
DT4.RecordSource = "select distinct 单号 from kpd where 锅号 in(select distinct 锅号 from pld where 编号='" & DH & "')"
DT4.Refresh
If Not DT4.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(3, 10) = "合同号"
        Excelapp.ActiveSheet.Cells(4, 10) = DT4.Recordset.Fields(0)
End If
        

L = 9
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 10
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方用量"   '配方用量
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "配方单位"       '单位
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "比例"       '比例
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "单位"    '配方单位
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With

L = 11
        
Do While Not DT1.Recordset.EOF

If L > 11 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT 浴比,染化助名称,配料用量,配料单位,配方,配方单位,校正值,车速,批次,染化助库  FROM pldd WHERE 料单编号='" & DH & "' and 工序名称='" & DT1.Recordset.Fields(0) & "'  order BY 次序号"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         'If InStr(dt2.Recordset.Fields(9), "助剂") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
        ' Else
        ' Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.name = "黑体"
        ' Excelapp.ActiveSheet.Range(25 * i + L, 6).Font.Bold = wdToggle
        ' Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.Size = 14
       '  Excelapp.ActiveSheet.Cells(25 * i + L, 6).Value = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
        ' End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         If InStr(dt2.Recordset.Fields(9), "助剂") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00000")
         End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '加网格线
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
'L = L + 1
DT1.Recordset.MoveNext
Loop


L = L + 1
'With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
'Excelapp.Selection.Merge
'Excelapp.Selection.Value = "备注："
'Excelapp.Selection.Borders.LineStyle = xlContinuous
'End With
If Len(Trim(bz)) > 2 Then
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 3, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = bz
Excelapp.Selection.Font.Size = 8
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If

L = L + 1

DT4.RecordSource = "select * from bgxx where 配料编号='" & DH & "'"
DT4.Refresh
If Not DT4.Recordset.EOF Then
DT4.Recordset.MoveFirst
bhxx = ""
Do While Not DT4.Recordset.EOF
bhxx = bhxx + DT4.Recordset.Fields(1) + "/"
DT4.Recordset.MoveNext
Loop
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = "并锅信息：" + Trim(bhxx)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If



        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

Set Excelapp = Nothing
Excelapp.Quit


End Sub


Public Sub plda(DT1 As Data, dt2 As Data, dt3 As Data, DH As String, DT4 As Adodc) ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim bhxx As String
        Dim gh As String


        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\pld.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
dt3.RecordSource = "SELECT count(*)  FROM plda WHERE 料单编号='" & DH & "'"
dt3.Refresh
JL = dt3.Recordset.Fields(0)  ''''''''''''''''''记录数
dt3.RecordSource = "SELECT DISTINCT 工序名称  FROM plda WHERE 料单编号='" & DH & "'"
dt3.Refresh
If Not dt3.Recordset.EOF Then
JL = JL + dt3.Recordset.RecordCount - 1
End If
If Int(JL / 19) = JL / 19 Then
MN = JL / 19
Else
MN = JL / 19 + 1
End If
dt3.RecordSource = "SELECT 料单编号,锅号,生产信息,压力,生产类别 as 颜色,配方单 as 色号,染化助单价 as 车台,审核 as 客户,重量,配料打印员,审核确认  FROM plda WHERE 料单编号='" & DH & "'"
dt3.Refresh

DT1.RecordSource = "SELECT 工序名称,水量  FROM plda WHERE 料单编号='" & DH & "' group by 工序名称,水量 Order BY 工序名称"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
'        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 1, 2), Excelapp.ActiveSheet.Cells(1, 6)).Select
'        Excelapp.Selection.Merge
'        Excelapp.Selection.Value = "枣庄华派集团有限公司配料单"
'        Excelapp.Selection.Font.Bold = True
'        Excelapp.Selection.Font.Size = 16
'        End With
        
        Excelapp.ActiveSheet.Cells(25 * i + 1, 2) = "河北广兴服饰有限公司配料单"
        Excelapp.ActiveSheet.Cells(25 * i + 1, 7) = "*" + Trim(DH) + "J*"
        
        Excelapp.ActiveSheet.Cells(25 * i + 2, 2) = "信息"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 6) = "打印日期"
        Excelapp.ActiveSheet.Cells(25 * i + 2, 7) = Now
        Excelapp.ActiveSheet.Cells(25 * i + 2, 8) = "操作员"     '''''原班次
        Excelapp.ActiveSheet.Cells(25 * i + 2, 4) = "配料编号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 3) = "锅号"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 4) = "品名"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 5) = "车台"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 6) = "客户"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 7) = "重量/匹数"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 8) = "颜色"
        Excelapp.ActiveSheet.Cells(25 * i + 3, 9) = "色号"
        gh = dt3.Recordset.Fields(1)                 '''''锅号'
        Excelapp.ActiveSheet.Cells(25 * i + 2, 5) = dt3.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 3) = dt3.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(25 * i + 2, 9) = dt3.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 3) = dt3.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 4) = dt3.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 5) = dt3.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 6) = dt3.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 7) = Format(dt3.Recordset.Fields(8), "#0.0") + "/" + Format(dt3.Recordset.Fields(10), "#0.0")
        Excelapp.ActiveSheet.Cells(25 * i + 4, 8) = dt3.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(25 * i + 4, 9) = dt3.Recordset.Fields(5)
        
    
DT4.RecordSource = "select distinct 单号 from kpd where 锅号 in(select distinct 锅号 from pld where 编号='" & DH & "')"
DT4.Refresh
If Not DT4.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(3, 10) = "合同号"
        Excelapp.ActiveSheet.Cells(4, 10) = DT4.Recordset.Fields(0)
End If
   
'打上网格线
'    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 3, 2), Excelapp.Sheets(1).Cells(25 * i + 4, 9)).Select '加网格线
'    Excelapp.Selection.Borders.LineStyle = xlContinuous
'    End With


'        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + 3, 2), Excelapp.ActiveSheet.Cells(25 * i + 4, 2)).Select
'        Excelapp.Selection.Merge
'        Excelapp.Selection.Value = "流程卡"
'        Excelapp.Selection.Borders.LineStyle = xlContinuous
'        End With


        Excelapp.ActiveSheet.Cells(6, 2) = "单循环时："
        Excelapp.ActiveSheet.Cells(6, 4) = "喷嘴位置："
        Excelapp.ActiveSheet.Cells(6, 6) = "主    泵："
        Excelapp.ActiveSheet.Cells(7, 2) = "水洗倍数："
        Excelapp.ActiveSheet.Cells(7, 4) = "提布轮速："
        Excelapp.ActiveSheet.Cells(7, 6) = "吸 水 率："
       ' Excelapp.ActiveSheet.Cells(7, 7) = xs


L = 9
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + 4, 2), Excelapp.Sheets(1).Cells(25 * i + 5, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = 10
'        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方用量"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "配方单位"
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "浴比"
        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = "染化助名称"
        'Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "配方"
        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = "配方用量"   '配方用量
       ' Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "校值"
        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = "配方单位"       '单位
        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = "比例"       '比例
        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = "单位"    '配方单位
        
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = "工艺"
        End With


L = 11
        


Do While Not DT1.Recordset.EOF

If L > 11 Then
Excelapp.ActiveSheet.Cells(25 * i + L, 2) = "工序"
Excelapp.ActiveSheet.Cells(25 * i + L, 3) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(25 * i + L, 4) = DT1.Recordset.Fields(1)

    With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 2), Excelapp.Sheets(1).Cells(25 * i + L, 4)).Select '加网格线
    Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With


L = L + 1
'Excelapp.ActiveSheet.Cells(25 * i + L, 2) = DT2.Recordset.Fields(0)
End If

dt2.RecordSource = "SELECT 浴比,染化助名称,配料用量,配料单位,配方,配方单位,校正值,车速,批次,染化助库  FROM plda WHERE 料单编号='" & DH & "' and 工序名称='" & DT1.Recordset.Fields(0) & "'  order BY 次序号"
dt2.Refresh

Excelapp.ActiveSheet.Cells(25 * i + L, 2) = dt2.Recordset.Fields(0)

Do While Not dt2.Recordset.EOF
'        Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(DT2.Recordset.Fields(1))
'        Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(DT2.Recordset.Fields(2), "#0.00000")
'        Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(3)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 6) = DT2.Recordset.Fields(4)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 7) = DT2.Recordset.Fields(5)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 8) = DT2.Recordset.Fields(6)
'        Excelapp.ActiveSheet.Cells(25 * i + L, 9) = DT2.Recordset.Fields(7)
         Excelapp.ActiveSheet.Cells(25 * i + L, 3) = Trim(dt2.Recordset.Fields(1))
         If InStr(dt2.Recordset.Fields(9), "助剂") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 6) = Format(Trim(dt2.Recordset.Fields(2)), "#0.0")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.name = "黑体"
         Excelapp.ActiveSheet.Range(25 * i + L, 6).Font.Bold = wdToggle
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).Font.Size = 14
         Excelapp.ActiveSheet.Cells(25 * i + L, 6).value = Format(Trim(dt2.Recordset.Fields(2)), "#0.000")
         End If
      '   Excelapp.ActiveSheet.Cells(25 * i + L, 5) = DT2.Recordset.Fields(6)
         Excelapp.ActiveSheet.Cells(25 * i + L, 5) = dt2.Recordset.Fields(5)
         If InStr(dt2.Recordset.Fields(9), "助剂") > 0 Then
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.0")
         Else
         Excelapp.ActiveSheet.Cells(25 * i + L, 4) = Format(dt2.Recordset.Fields(4) * dt2.Recordset.Fields(6), "#0.00000")
         End If
         Excelapp.ActiveSheet.Cells(25 * i + L, 7) = dt2.Recordset.Fields(3)
         
        With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(25 * i + L, 8), Excelapp.ActiveSheet.Cells(25 * i + L, 9)).Select
        Excelapp.Selection.Merge
        Excelapp.Selection.value = Trim(dt2.Recordset.Fields(7))
        End With
         
    
         With Excelapp.Sheets(1).Range(Excelapp.Sheets(1).Cells(25 * i + L, 3), Excelapp.Sheets(1).Cells(25 * i + L, 7)).Select '加网格线
         Excelapp.Selection.Borders(xlEdgeBottom).LineStyle = xlDot
         End With
         
L = L + 1
dt2.Recordset.MoveNext
Loop
'Excelapp.ActiveShee.PageSetup.PrintGridlines = True
L = L + 1
DT1.Recordset.MoveNext
Loop


L = L + 2
'With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 2), Excelapp.ActiveSheet.Cells(L + 1, 2)).Select
'Excelapp.Selection.Merge
'Excelapp.Selection.Value = "备注："
'Excelapp.Selection.Borders.LineStyle = xlContinuous
'End With
DT4.RecordSource = "select 总备注 from sczy_z  where 单号 in(select distinct 单号 from kpd where 锅号='" & gh & "' and len(isnull(单号,0))>0)"
DT4.Refresh


If Not DT4.Recordset.EOF Then
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 3, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = DT4.Recordset.Fields(0)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If

L = L + 5

DT4.RecordSource = "select * from bgxx where 配料编号='" & DH & "'"
DT4.Refresh
If Not DT4.Recordset.EOF Then
DT4.Recordset.MoveFirst
bhxx = ""
Do While Not DT4.Recordset.EOF
bhxx = bhxx + DT4.Recordset.Fields(1) + "/"
DT4.Recordset.MoveNext
Loop
With Excelapp.ActiveSheet.Range(Excelapp.ActiveSheet.Cells(L, 3), Excelapp.ActiveSheet.Cells(L + 1, 9)).Select
Excelapp.Selection.Merge
Excelapp.Selection.WrapText = True
Excelapp.Selection.value = "并锅信息：" + Trim(bhxx)
Excelapp.Selection.Borders.LineStyle = xlContinuous
End With
End If
'Next
'Excelapp.ActiveSheet.PrintOut To:=MN
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'        Excelapp.Quit
'Set Excelapp = Nothing
'        Exit Sub

'Ert:

'Excelapp.Quit '关闭EXCEL
'Excelapp.ActiveWorkbook.Saved = True
'Excelapp.Workbooks.Close
'Excelapp.Quit
'Set Excelapp = Nothing
'''''''''''''''''''''''''''''''''''''
        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub




