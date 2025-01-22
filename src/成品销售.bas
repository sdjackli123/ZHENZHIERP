Attribute VB_Name = "成品销售"
Public Sub dzmx(Flex As VSFlexGrid, fd1, fd2, BT)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\dzmx.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k + 1

                          For j = 1 To .Cols

                              
                              DoEvents
                              Excelapp.ActiveSheet.Cells(i + 2, j) = "'" & .TextMatrix(i - 1, j)
                              If i >= 2 And (j = (fd1 - 1) Or j = (fd2 - 1)) Then
                              Excelapp.ActiveSheet.Cells(i + 2, j) = Val(Excelapp.ActiveSheet.Cells(i + 2, j))
                              End If
                          Next j
               
         If i >= 2 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i + 2, fd1 - 1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i + 2, fd2 - 1)) + Q2
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(2, 2) = BT + "  客户对账单"

Excelapp.ActiveSheet.Cells(i + 1, 1) = "合计"
Excelapp.ActiveSheet.Cells(i + 1, fd1 - 1) = Q1
Excelapp.ActiveSheet.Cells(i + 1, fd2 - 1) = Q2

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



Public Sub OutadodcToExcelBC(Flex As VSFlexGrid, FD, BT)

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
               
         If i >= 3 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, FD) = Q

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

Public Sub CPCKOutadodcToExcel(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ps As Integer, sl As Single) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品发货.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开,isnull(光坯,0)  FROM v_jgmx WHERE 单号='" & DH & "' and 顺序号 between '" & xh1 & "' and '" & xh2 & "' order BY 顺序号"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '''客户名称
        Excelapp.ActiveSheet.Cells(2, 6) = Trim(DT1.Recordset.Fields(9))   '''日期
        Excelapp.ActiveSheet.Cells(2, 10) = DH    '''单号
        Excelapp.ActiveSheet.Cells(9, 7) = ps   '''匹数
        Excelapp.ActiveSheet.Cells(9, 8) = sl   '''数量

i = 4
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    ''''锅号
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)    ''''款号
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(2)    ''''颜色
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(5)    ''''匹数
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(6)    ''''数量
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(13)    '''''光坯
        Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(10)   ''''备注
i = i + 1
DT1.Recordset.MoveNext
Loop


DT1.RecordSource = "SELECT round(sum(isnull(光坯,0)),2)  FROM v_jgmx WHERE 单号='" & DH & "'"
DT1.Refresh
If Not IsNull(DT1.Recordset.Fields(0)) Then
        Excelapp.ActiveSheet.Cells(9, 9) = DT1.Recordset.Fields(0)   '''光坯重量
End If

DT1.RecordSource = "SELECT 模块  FROM yhb WHERE 用户='" & yhm & "'"
DT1.Refresh
If Not IsNull(DT1.Recordset.Fields(0)) Then
    Excelapp.ActiveSheet.Cells(11, 10) = Trim(DT1.Recordset.Fields(0))  '''用户模板
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
Set Excelapp = Nothing
Excelapp.Quit
End Sub



Public Sub CPCKTZD(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品结算.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT 客户全称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开 as 色号,跟单,isnull(光坯,0),负责,计划号,isnull(米数,0),单号,'' as 下单日期,核算  FROM v_jgmx WHERE 单号='" & DH & "' and 顺序号 between '" & xh1 & "' and '" & xh2 & "' order BY 顺序号"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''客户
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''日期
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''单据
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''负责
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''款号
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''缸号
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''颜色
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''匹数
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''毛胚重量
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''光坯重量
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''单价
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''金额
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''备注
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'dt3.RecordSource = "SELECT SUM(ISNULL(匹数,0)),SUM(ISNULL(数量,0)),SUM(ISNULL(光坯,0)),SUM(ISNULL(金额,0))  FROM v_jgmx WHERE 单号='" & DH & "' and (加工类别='成品布' or 加工类别='染色费' or 加工类别='定型费' or 加工类别='不收费' or 加工类别='外印花' or 加工类别='只磨毛')"
'dt3.Refresh
'If Not dt3.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(10, 1) = "页" + Trim(ym)  ''''页号
'        Excelapp.ActiveSheet.Cells(10, 3) = "总页" + Trim(ys)  ''''页数
'        Excelapp.ActiveSheet.Cells(10, 4) = "合计"    ''''品名
'        Excelapp.ActiveSheet.Cells(10, 8) = Format(dt3.Recordset.Fields(0), "#0.0")   ''''匹数
'        Excelapp.ActiveSheet.Cells(10, 9) = Format(dt3.Recordset.Fields(1), "#0.00")   ''''毛坯数量
'        Excelapp.ActiveSheet.Cells(10, 10) = Format(dt3.Recordset.Fields(2), "#0.00")   ''''光坯数量
'End If

'dt3.RecordSource = "SELECT SUM(ISNULL(金额,0))  FROM v_jgmx WHERE 单号='" & DH & "'"
'dt3.Refresh
'If Not dt3.Recordset.EOF Then
'        Excelapp.ActiveSheet.Cells(10, 11) = Format(dt3.Recordset.Fields(0), "#0.00") ''''金额
'End If

Excelapp.ActiveWindow.Zoom = 100
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
Public Sub CPCKTZDGP(DT1 As Adodc, dt3 As Adodc, DT4 As Adodc, dt5 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer)

    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    Dim x   As Integer
    On Error GoTo Ert

    ' 创建Excel应用程序实例
    Dim Excelapp   As Excel.Application
    Set Excelapp = New Excel.Application

    On Error Resume Next

    ' 打开指定的工作簿
    Excelapp.SheetsInNewWorkbook = 1
    Excelapp.Caption = "广兴染整软件之打印"
    Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品结算光坯.xls")

    ' 激活第一个工作表
    Excelapp.Sheets(1).Activate

    ' 加载数据源
    DT1.RecordSource = "SELECT 客户全称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开 as 色号,跟单,isnull(光坯,0),负责,计划号,isnull(米数,0),单号,'' as 下单日期,核算,技术要求,来料单位,业务  FROM v_jgmx WHERE 单号='" & DH & "' and 顺序号 between '" & xh1 & "' and '" & xh2 & "' order BY 顺序号"
    DT1.Refresh
    dt5.RecordSource = "SELECT * FROM yskzcx where 客户= '" & DT1.Recordset.Fields(0).value & "'"
    dt5.Refresh
    DT4.RecordSource = "SELECT round(sum(isnull(欠款,0)),2) as 合计欠款 FROM jgzcx where 客户= '" & DT1.Recordset.Fields(0).value & "'"
    DT4.Refresh

    ' 填充Excel数据
    If Not DT1.Recordset.EOF Then
        DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   ' 客户
        Excelapp.ActiveSheet.Cells(3, 13) = Trim(DT1.Recordset.Fields(9))  ' 日期
        Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(18)       ' 单据
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)       ' 负责
        Excelapp.ActiveSheet.Cells(3, 7) = DT4.Recordset.Fields(0) & "元"  ' 累计欠款
        
        i = 5
        Do While Not DT1.Recordset.EOF
            Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ' 品名
            Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ' 缸号
            Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)    ' 颜色
            Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ' 匹数
            Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ' 毛胚重量
            Excelapp.ActiveSheet.Cells(i, 9) = Val(DT1.Recordset.Fields(21)) ' 克重
            Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(7)   ' 单价
            Excelapp.ActiveSheet.Cells(i, 11) = Val(DT1.Recordset.Fields(8)) ' 金额
            Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(22)  ' 来料
            Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(10)  ' 备注
            Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(13)  ' 业务
            i = i + 1
            DT1.Recordset.MoveNext
        Loop
    End If

    ' 显示Excel并允许用户进行编辑
    Excelapp.Visible = True
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.DisplayAlerts = False
    
    ' 弹出打印预览
    Excelapp.ActiveSheet.PrintPreview
    
    ' 打印后退出
    Excelapp.Quit
    Set Excelapp = Nothing
    
    Exit Sub

Ert:
    ' 错误处理
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
End Sub


Public Sub CPCKTZDF(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品结算.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT 客户全称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开 as 色号,跟单,isnull(光坯,0),负责,计划号,isnull(米数,0),单号,'' as 下单日期,核算  FROM v_jgmx WHERE 单号='" & DH & "' and 顺序号 between '" & xh1 & "' and '" & xh2 & "'  order BY 顺序号"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''客户
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''日期
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''单据
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''负责
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''款号
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''缸号
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''颜色
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''匹数
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''毛胚重量
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''光坯重量
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''单价
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''金额
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''备注
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If
'        Excelapp.ActiveSheet.Cells(10, 1) = "页" + Trim(ym)  ''''页号
'        Excelapp.ActiveSheet.Cells(10, 3) = "总页" + Trim(ys)  ''''页数

Excelapp.ActiveWindow.Zoom = 100
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

Public Sub CPCKTZDFGP(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ym As Integer, ys As Integer) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品结算光坯.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        

DT1.RecordSource = "SELECT 客户全称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开 as 色号,跟单,isnull(光坯,0),负责,计划号,isnull(米数,0),单号,'' as 下单日期,核算  FROM v_jgmx WHERE 单号='" & DH & "' and 顺序号 between '" & xh1 & "' and '" & xh2 & "'  order BY 顺序号"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = Trim(DT1.Recordset.Fields(0))   '''客户
        Excelapp.ActiveSheet.Cells(3, 12) = Trim(DT1.Recordset.Fields(9))    '''日期
        Excelapp.ActiveSheet.Cells(2, 12) = DT1.Recordset.Fields(18)    '''单据
        Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(15)    '''负责
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(4)    ''''款号
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)    ''''缸号
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2) ''''''''''''''+ DT1.Recordset.Fields(12)   ''''颜色
        Excelapp.ActiveSheet.Cells(i, 7) = Val(DT1.Recordset.Fields(5)) ''''匹数
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(6)) ''''毛胚重量
        Excelapp.ActiveSheet.Cells(i, 13) = Val(DT1.Recordset.Fields(14))    '''''光坯重量
      '  Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    '''''单价
      '  Excelapp.ActiveSheet.Cells(i, 12) = Val(DT1.Recordset.Fields(8))    '''''金额
        Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(10)    '''''备注
        
i = i + 1
DT1.Recordset.MoveNext
Loop
End If
'        Excelapp.ActiveSheet.Cells(10, 1) = "页" + Trim(ym)  ''''页号
'        Excelapp.ActiveSheet.Cells(10, 3) = "总页" + Trim(ys)  ''''页数

Excelapp.ActiveWindow.Zoom = 100
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

Public Sub CLRKOutadodcToExcel(DT1 As Adodc, DH As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\CLRK.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(2).Activate
DT1.RecordSource = "SELECT 供应单位,材料名称,材料规格,材料单位,颜色,批次,数量,单价,合计金额,日期  FROM CLGL WHERE 单据号='" & DH & "' order BY 序号"
DT1.Refresh
DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 5) = DH
        Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(9)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(8)
        
i = i + 1
DT1.Recordset.MoveNext
Loop
DT1.RecordSource = "SELECT *  FROM CKGL WHERE 单据号='" & DH & "' order BY 序号"
DT1.Refresh


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

Public Sub CPCKQKT(DT1 As Adodc, DH As String, je As Single) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\欠款条.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        
DT1.RecordSource = "SELECT 客户名称,品名,颜色,锅号,和约号,匹数,数量,单价,金额,日期,备注,加工类别,发票已开  FROM v_jgmx WHERE 单号='" & DH & "'"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(4, 6) = DH  '''单号
        Excelapp.ActiveSheet.Cells(6, 6) = Format(je, "#0.00")  '''金额
         Excelapp.ActiveSheet.Cells(9, 6) = DT1.Recordset.Fields(0) '''客户
       Excelapp.ActiveSheet.Cells(14, 6) = Trim(DT1.Recordset.Fields(9))  '''日期

Excelapp.ActiveWindow.Zoom = 100


        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

Excelapp.Quit
Set Excelapp = Nothing
End Sub


Public Sub fhdmd(DT1 As Adodc, gh As String, DH As String, dj As String)   ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户全称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区  FROM v_bmd WHERE 锅号='" & gh & "' and 缸号='" & DH & "' and 单据='" & dj & "' order by 匹号"
DT1.Refresh

If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 7) = Trim(DT1.Recordset.Fields(10))
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(7)
i = 1
L = 0
Do While Not DT1.Recordset.EOF
If Int(i / 19) = i / 19 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select  重量,count(匹号) as 匹数 from v_bmd where 锅号='" & gh & "' and 缸号='" & DH & "' and 单据='" & dj & "' group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
DT1.Recordset.MoveNext
Loop
End If

'Excelapp.ActiveSheet.Cells(32, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps


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

Public Sub CPCKSH(DT1 As Adodc, DH As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\成品出库审核OK.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT 客户,品名,颜色,锅号,和约号 as 客户合同号,匹数,数量,单价,金额,日期,isnull(光坯,0),计划号,幅宽+'*'+克重,isnull(米数,0),负责,核算  FROM v_jgmx WHERE 单号='" & DH & "'  order BY 顺序号"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)   '''购货单位
        Excelapp.ActiveSheet.Cells(2, 5) = Date   '''日期
        Excelapp.ActiveSheet.Cells(2, 9) = DH    '''单号
        Excelapp.ActiveSheet.Cells(17, 2) = DT1.Recordset.Fields(14)   '''业务
        
i = 5
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(12)  ''''幅宽
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)    ''''颜色
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(5)    ''''匹数
        If Val(DT1.Recordset.Fields(13)) <> 0 Then
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(13)   ''''米数
        Else
        Excelapp.ActiveSheet.Cells(i, 5) = ""  ''''米数
        End If
        If DT1.Recordset.Fields(20) = "毛坯" Then
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(6), "#0.00")   ''''毛坯重量
        Else
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(10), "#0.00")   ''''光坯重量
        End If
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(7)    '''''单价
        Excelapp.ActiveSheet.Cells(i, 8) = Val(DT1.Recordset.Fields(8))    '''''金额
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(11) ''''计划号
        Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(4)    '''''客户合同号
i = i + 1
DT1.Recordset.MoveNext
Loop


DT1.RecordSource = "SELECT SUM(ISNULL(匹数,0)),SUM(ISNULL(数量,0)),SUM(ISNULL(光坯,0)),SUM(ISNULL(金额,0)),round(sum(isnull(米数,0)),1)  FROM v_jgmx WHERE 单号='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(14, 1) = "合计"    ''''品名
        Excelapp.ActiveSheet.Cells(14, 4) = Format(DT1.Recordset.Fields(0), "#0.0")   ''''匹数
        Excelapp.ActiveSheet.Cells(14, 5) = Format(DT1.Recordset.Fields(4), "#0.0")   ''''匹数
        Excelapp.ActiveSheet.Cells(14, 6) = Format(DT1.Recordset.Fields(2), "#0.00")   ''''毛坯数量
        Excelapp.ActiveSheet.Cells(14, 8) = Format(DT1.Recordset.Fields(3), "#0.00") ''''金额
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

Public Sub rkdmd(DT1 As Adodc, gh As String, dj As String, xh As String, DH As String, xh1 As Integer)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\rkmd.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,成分,匹号,日期,合同负责,下单日期,码数,单号  FROM v_bmd WHERE 锅号='" & gh & "' and 入库单据='" & dj & "' and 入库序号='" & xh1 & "'  and 单号='" & DH & "' and 序号='" & xh & "' order by 匹号"
DT1.Refresh

If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(14)      '''单号
        Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(11)    '''合同负责
        Excelapp.ActiveSheet.Cells(5, 3) = Trim(DT1.Recordset.Fields(10))  '''入库日期
        Excelapp.ActiveSheet.Cells(5, 9) = Trim(DT1.Recordset.Fields(12))  '''下单日期
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0)         '''客户
        Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(3)         '''品种
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2)         '''款号
        Excelapp.ActiveSheet.Cells(8, 5) = DT1.Recordset.Fields(1)         ''锅号
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(8)    ''''成分
        Excelapp.ActiveSheet.Cells(8, 11) = DT1.Recordset.Fields(5)    '''色别
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4)    '''幅宽
        Excelapp.ActiveSheet.Cells(10, 5) = DT1.Recordset.Fields(7)    '''克重
i = 1
L = 0
mpps = 0
mpzl = 0
Do While Not DT1.Recordset.EOF
If Int(i / 19) = i / 19 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = Val(DT1.Recordset.Fields(13))
mpps = mpps + 1
mpzl = mpzl + Val(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

Excelapp.ActiveSheet.Cells(10, 8) = mpps

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

