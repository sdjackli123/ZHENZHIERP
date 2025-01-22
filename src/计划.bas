Attribute VB_Name = "计划"
Public Sub jhbOutadodcToExcel(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\jhb.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j - 1)
                      
                          Next j
               
         If i >= 1 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
Excelapp.ActiveSheet.Cells(1, 1) = BT
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

Public Sub jh3(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\印花流程单条码.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select max(印花数量) as zl from v_yhjh where 印花锅号='" & gh & "' and 版号 is not null and len(版号)>0"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from v_yhjh where 印花锅号='" & gh & "' And 印花数量 = '" & a & "' and 版号 is not null and len(版号)>0"
DT1.Refresh
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '客户
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))   '''款号
Excelapp.ActiveSheet.Cells(2, 10) = Trim(DT1.Recordset.Fields(23))   ''''日期
Excelapp.ActiveSheet.Cells(2, 13) = Trim(DT1.Recordset.Fields(3))      ''''''锅号
'''Excelapp.ActiveSheet.Cells(3, 2) = dt1.Recordset.Fields(13)        ''''
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(5)     '''品名
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(4)    ''色别
Excelapp.ActiveSheet.Cells(5, 3) = Trim(gh)                           ''印花锅号
Excelapp.ActiveSheet.Cells(3, 4) = DT1.Recordset.Fields(24)    '印花款号
Excelapp.ActiveSheet.Cells(5, 9) = "*" + Trim(DT1.Recordset.Fields(3)) + "J*"            ''' 染色锅号条码

DT1.RecordSource = "select round(sum(印花数量),2),sum(印花匹数) from v_yhjh where 印花锅号='" & gh & "' and 版号 is not null and len(版号)>0"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Excelapp.ActiveSheet.Cells(4 + 22, 9) = DT1.Recordset.Fields(0)
End If

DT1.RecordSource = "select * from v_yhjh where 印花锅号='" & gh & "' and 版号 is not null and len(版号)>0"
DT1.Refresh
i = 0
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(10 + i * 1, 1) = DT1.Recordset.Fields(5)    '''品名
Excelapp.ActiveSheet.Cells(10 + i * 1, 4) = DT1.Recordset.Fields(21)   ''
Excelapp.ActiveSheet.Cells(10 + i * 1, 6) = DT1.Recordset.Fields(20)
Excelapp.ActiveSheet.Cells(10 + i * 1, 7) = DT1.Recordset.Fields(19)  ''库位
Excelapp.ActiveSheet.Cells(10 + i * 1, 8) = DT1.Recordset.Fields(17) '''''''''版号
Excelapp.ActiveSheet.Cells(10 + i * 1, 9) = DT1.Recordset.Fields(22)  '''备注
Excelapp.ActiveSheet.Cells(10 + i * 1, 13) = DT1.Recordset.Fields(18) '''''图案
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Public Sub jh33(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\印花流程单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select distinct 印花锅号,客户名称,CONVERT(varchar,计划日期, 23),款号,印花款号 from v_yhjh where 印花单号='" & gh & "' and 版号 is not null and len(版号)>0 order by 印花锅号"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(2, 14) = Trim(gh)
Excelapp.ActiveSheet.Cells(2, 16) = DT1.Recordset.Fields(4)

i = 4
Do While Not DT1.Recordset.EOF

dt2.RecordSource = "select * from v_yhjh where 印花锅号='" & DT1.Recordset.Fields(0) & "' and 版号 is not null and len(版号)>0"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(5)    '''品名
Excelapp.ActiveSheet.Cells(i, 4) = dt2.Recordset.Fields(3)     ''''锅号
Excelapp.ActiveSheet.Cells(i, 5) = dt2.Recordset.Fields(4)     ''色别
Excelapp.ActiveSheet.Cells(i, 6) = Trim(dt2.Recordset.Fields(21))  '匹数
Excelapp.ActiveSheet.Cells(i, 7) = Trim(dt2.Recordset.Fields(20))  ''重量
Excelapp.ActiveSheet.Cells(i, 8) = dt2.Recordset.Fields(19)  ''库位
Excelapp.ActiveSheet.Cells(i, 9) = dt2.Recordset.Fields(16)  ''编号
Excelapp.ActiveSheet.Cells(i, 10) = dt2.Recordset.Fields(22)    '''备注
Excelapp.ActiveSheet.Cells(i, 14) = dt2.Recordset.Fields(18) '''''图案
Excelapp.ActiveSheet.Cells(i, 16) = dt2.Recordset.Fields(17)  '''''''''版号
i = i + 1
dt2.Recordset.MoveNext
Loop

DT1.Recordset.MoveNext
Loop
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


Public Sub pcjh(DT1 As Adodc, dt2 As Adodc, sql1 As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\染色计划.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.RecordSource = "SELECT 车台编号 FROM CT ORDER BY ip"
DT1.Refresh


If Not DT1.Recordset.EOF Then
i = 4
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
dt2.RecordSource = "SELECT top 6 * FROM v_kpdb where  (" + sql1 + ") and 车台='" & DT1.Recordset.Fields(0) & "' ORDER BY 排产编号"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
L = 1
Do While Not dt2.Recordset.EOF
'Excelapp.ActiveSheet.Cells(i, 2) = dt2.Recordset.Fields(0)    '''品名
Excelapp.ActiveSheet.Cells(i, 3) = dt2.Recordset.Fields(1)     ''''排产时间
Excelapp.ActiveSheet.Cells(i, 4) = dt2.Recordset.Fields(2)     ''排产编号
Excelapp.ActiveSheet.Cells(i, 5) = dt2.Recordset.Fields(3)  '客户名称
Excelapp.ActiveSheet.Cells(i, 6) = dt2.Recordset.Fields(4)  ''品名
Excelapp.ActiveSheet.Cells(i, 7) = dt2.Recordset.Fields(5)  ''色号
Excelapp.ActiveSheet.Cells(i, 8) = dt2.Recordset.Fields(6)  ''颜色
Excelapp.ActiveSheet.Cells(i, 9) = dt2.Recordset.Fields(7)    '''重量
Excelapp.ActiveSheet.Cells(i, 10) = dt2.Recordset.Fields(8) ''''锅号
Excelapp.ActiveSheet.Cells(i, 11) = dt2.Recordset.Fields(9)  '''''''''排产备注
Excelapp.ActiveSheet.Cells(i, 12) = dt2.Recordset.Fields(9)  '''''''''操作
i = i + 1
dt2.Recordset.MoveNext
L = L + 1
Loop
i = i + 7 - L
Else
i = i + 6
End If

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

Public Sub pgk(DT1 As Adodc, gh As String, selectedPrinter As String)
    Dim Excelapp As Object  ' 声明 Excel 应用程序对象
    On Error Resume Next    ' 出现错误时继续执行下一行

    Set Excelapp = CreateObject("Excel.Application")  ' 创建 Excel 应用程序对象

    If Excelapp Is Nothing Then   ' 如果 Excel 应用程序对象未创建成功
        MsgBox "Excel is not installed on this machine."   ' 显示错误信息
        Exit Sub   ' 退出子程序
    End If

    On Error GoTo Ert   ' 出现错误时转到 Ert 标签

    Excelapp.Caption = "广兴打印模版软件之打印"   ' 设置 Excel 窗口标题
    Excelapp.SheetsInNewWorkbook = 1   ' 设置新工作簿中的工作表数为 1
    Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\排缸卡.xls")   ' 打开已存在的工作簿
    Excelapp.Sheets(1).Activate   ' 激活第一个工作表

    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"   ' 设置数据库连接字符串
    DT1.RecordSource = "select 客户名称,锅号,品名,标签,色别,匹数,重量,dr,日期,色名,技术要求,备注 from kpd where 锅号='" & gh & "' "   ' 设置查询语句
    DT1.Refresh   ' 刷新数据

    If Not DT1.Recordset.EOF Then   ' 如果记录集不为空
        DT1.Recordset.MoveFirst   ' 移动到记录集的第一条记录
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   ' 将客户名称填写到单元格
        Excelapp.ActiveSheet.Cells(3, 4) = Trim(DT1.Recordset.Fields(1))   ' 将锅号填写到单元格
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(2)   ' 将品名填写到单元格
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)   ' 将颜色填写到单元格
        Excelapp.ActiveSheet.Cells(5, 4) = DT1.Recordset.Fields(9)   ' 将色号填写到单元格
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(5)   ' 将匹数填写到单元格
        Excelapp.ActiveSheet.Cells(6, 4) = DT1.Recordset.Fields(6)   ' 将重量填写到单元格
    End If

    Excelapp.ActiveWindow.Zoom = 100   ' 设置窗口缩放比例为 100%
    ' ExcelApp.Visible = True  ' 注释掉设置 Excel 应用程序可见的代码
    Excelapp.DisplayAlerts = False   ' 禁用显示警告

    ' 切换到用户选择的打印机
    If selectedPrinter <> "" Then
        TrySetActivePrinter Excelapp, selectedPrinter
    End If

    ' 打印工作表
    Excelapp.ActiveSheet.PrintOut Copies:=1, Preview:=False, PrintToFile:=False, Collate:=True

Cleanup:
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
    Exit Sub

Ert:   ' 错误处理标签
    MsgBox "An error occurred: " & Err.Description   ' 显示错误信息
    Excelapp.Quit   ' 退出 Excel 应用程序
    Set Excelapp = Nothing   ' 释放 Excel 应用程序对象
End Sub
Public Sub pgk1(DT1 As Adodc, gh As String)
    Dim Excelapp As Object  ' 声明 Excel 应用程序对象
    On Error Resume Next    ' 出现错误时继续执行下一行

    Set Excelapp = CreateObject("Excel.Application")  ' 创建 Excel 应用程序对象

    If Excelapp Is Nothing Then   ' 如果 Excel 应用程序对象未创建成功
        MsgBox "Excel is not installed on this machine."   ' 显示错误信息
        Exit Sub   ' 退出子程序
    End If

    On Error GoTo Ert   ' 出现错误时转到 Ert 标签

    Excelapp.Caption = "广兴打印模版软件之打印"   ' 设置 Excel 窗口标题
    Excelapp.SheetsInNewWorkbook = 1   ' 设置新工作簿中的工作表数为 1
    Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\排缸卡.xls")   ' 打开已存在的工作簿
    Excelapp.Sheets(1).Activate   ' 激活第一个工作表

    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"   ' 设置数据库连接字符串
    DT1.RecordSource = "select 客户名称,锅号,品名,标签,色别,匹数,重量,dr,日期,色名,技术要求,备注 from kpd where 锅号='" & gh & "' "   ' 设置查询语句
    DT1.Refresh   ' 刷新数据

    If Not DT1.Recordset.EOF Then   ' 如果记录集不为空
        DT1.Recordset.MoveFirst   ' 移动到记录集的第一条记录
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   ' 将客户名称填写到单元格
        Excelapp.ActiveSheet.Cells(3, 4) = Trim(DT1.Recordset.Fields(1))   ' 将锅号填写到单元格
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(2)   ' 将品名填写到单元格
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)   ' 将颜色填写到单元格
        Excelapp.ActiveSheet.Cells(5, 4) = DT1.Recordset.Fields(9)   ' 将色号填写到单元格
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(5)   ' 将匹数填写到单元格
        Excelapp.ActiveSheet.Cells(6, 4) = DT1.Recordset.Fields(6)   ' 将重量填写到单元格
    End If

    Excelapp.ActiveWindow.Zoom = 100   ' 设置窗口缩放比例为 100%
    ' Excelapp.Visible = True  ' 注释掉设置 Excel 应用程序可见的代码
    Excelapp.DisplayAlerts = False   ' 禁用显示警告

    Excelapp.ActiveSheet.PrintOut   ' 直接打印当前工作表

    Set Excelapp = Nothing   ' 释放 Excel 应用程序对象
    Exit Sub   ' 退出子程序

Ert:   ' 错误处理标签
    MsgBox "An error occurred: " & Err.Description   ' 显示错误信息
    Excelapp.Quit   ' 退出 Excel 应用程序
    Set Excelapp = Nothing   ' 释放 Excel 应用程序对象
End Sub

Private Sub TrySetActivePrinter(ByRef Excelapp As Object, ByVal PrinterName As String)
    On Error Resume Next
    Dim CurrentPrinter As String
    CurrentPrinter = Excelapp.ActivePrinter
    Excelapp.ActivePrinter = PrinterName
    If Err.Number <> 0 Then
        ' 尝试附加端口名称
        Excelapp.ActivePrinter = PrinterName & " on " & Split(PrinterName, " (")(1) ' 提取并附加端口名称
        If Err.Number = 0 Then
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub pcmx(Flex As VSFlexGrid, BT As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\pcmx.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q = 0

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
               
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT


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

