Attribute VB_Name = "纺织"
Public Sub ddlcd(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴针织软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\纺织\广兴\生产计划单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,款号,织号,品名,筒颈,计划,克重,幅宽,颜色,开幅线,备注,交期,日期,序号,纱别,车间,匹重 FROM v_kpd_ddjh where 单据='" & Zh & "' order by 序号"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''织号
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''客户
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''车间
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''款号
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4品名
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''颜色
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''寸数
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''计划

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''克重
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''幅宽
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''匹重
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''开幅线
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''交期
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''备注


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT 织号,sum(计划) FROM v_kpd_ddjh where 单据='" & Zh & "' group by 织号 order by 织号"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT 织号,纱支,批次,织耗,配比,备注,颜色,产地 FROM sxpb where 织号='" & DT1.Recordset.Fields(0) & "' order by 序号"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''织号
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''纱支
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''颜色
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''产地
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''批次
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''织耗
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''配比
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''纱量
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''备注
L = L + 1
dt2.Recordset.MoveNext
Loop
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

Public Sub ddlcdxz(DT1 As Adodc, dt2 As Adodc, Zh As String, fw As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴针织软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\纺织\广兴\生产计划单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,款号,织号,品名,筒颈,计划,克重,幅宽,颜色,开幅线,备注,交期,日期,序号,纱别,车间,匹重 FROM v_kpd_ddjh where  织号 in(" + fw + ") order by 织号,序号"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''织号
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''客户
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''车间
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''款号
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4品名
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''颜色
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''寸数
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''计划

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''克重
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''幅宽
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''匹重
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''开幅线
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''交期
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''备注


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT 织号,sum(计划) FROM v_kpd_ddjh where  织号 in(" + fw + ") group by 织号 order by 织号"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT 织号,纱支,批次,织耗,配比,备注,颜色,产地 FROM sxpb where 织号='" & DT1.Recordset.Fields(0) & "' order by 序号"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''织号
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''纱支
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''颜色
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''产地
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''批次
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''织耗
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''配比
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''纱量
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''备注
L = L + 1
dt2.Recordset.MoveNext
Loop
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

Public Sub ddlcdjh(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴针织软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\纺织\广兴\生产计划单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,款号,织号,品名,筒颈,计划,克重,幅宽,颜色,开幅线,备注,交期,日期,序号,纱别,车间,匹重 FROM v_kpd_ddjh_cjjt where 单据='" & Zh & "' order by 序号"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 3
L = 11
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = Trim(DT1.Recordset.Fields(2))  ''''织号
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(0)         '''客户
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(15)         '''车间
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''款号
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)    ''''''''4品名
Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)   '''''''颜色
Excelapp.ActiveSheet.Cells(i, 12) = DT1.Recordset.Fields(4)     '''''''''寸数
Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(5)    '''''计划

Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)    '''''克重
Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(7)    '''''幅宽
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(16)    ''''匹重
Excelapp.ActiveSheet.Cells(i, 17) = DT1.Recordset.Fields(9)   ''''''开幅线
Excelapp.ActiveSheet.Cells(i, 18) = Trim(DT1.Recordset.Fields(11))   ''''''交期
'Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)   ''''''备注


i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT 织号,sum(计划) FROM v_kpd_ddjh_cjjt where 单据='" & Zh & "' group by 织号 order by 织号"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT 织号,纱支,批次,织耗,配比,备注,颜色,产地 FROM sxpb where 织号='" & DT1.Recordset.Fields(0) & "' order by 序号"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(L, 1) = Trim(dt2.Recordset.Fields(0))  ''''织号
Excelapp.ActiveSheet.Cells(L, 3) = dt2.Recordset.Fields(1)         '''纱支
Excelapp.ActiveSheet.Cells(L, 7) = dt2.Recordset.Fields(6)                        ''''颜色
Excelapp.ActiveSheet.Cells(L, 9) = dt2.Recordset.Fields(7)                        '''产地
Excelapp.ActiveSheet.Cells(L, 11) = dt2.Recordset.Fields(2)                        ''''批次
Excelapp.ActiveSheet.Cells(L, 12) = Val(dt2.Recordset.Fields(3))    ''''''''织耗
Excelapp.ActiveSheet.Cells(L, 13) = Val(dt2.Recordset.Fields(4))    '''''''配比
Excelapp.ActiveSheet.Cells(L, 14) = Val(DT1.Recordset.Fields(1)) * 100 / (100 - Val(dt2.Recordset.Fields(3))) * Val(dt2.Recordset.Fields(4)) / 100 '''''''''纱量
Excelapp.ActiveSheet.Cells(L, 15) = dt2.Recordset.Fields(5)     '''''''''备注
L = L + 1
dt2.Recordset.MoveNext
Loop
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

Public Sub dbqww(DT1 As Adodc, dt2 As Adodc, DH As String, k As Long, bh As Long, cj As String, jh As String, bz As String)  ''''无标题

        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴针织软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\纺织\广兴\tmww.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,合同号,单号,织号,品名,幅宽,克重,纱别,对账,车台,筒颈,开幅线,颜色 FROM kpd WHERE 织号='" & DH & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst

L = 0
TM = Mid("00000000", 1, 8 - Len(Trim(bh))) + Trim(bh)

        Excelapp.ActiveSheet.Cells(L * 20 + 2, 2) = DH    '''''织号
        Excelapp.ActiveSheet.Cells(L * 20 + 3, 2) = DT1.Recordset.Fields(1)  '''''合同号
        Excelapp.ActiveSheet.Cells(L * 20 + 3, 4) = bz  '''''换批次
        Excelapp.ActiveSheet.Cells(L * 20 + 4, 2) = DT1.Recordset.Fields(0)  ''''客户
        Excelapp.ActiveSheet.Cells(L * 20 + 5, 2) = DT1.Recordset.Fields(4) '''品种
        
        
        Excelapp.ActiveSheet.Cells(L * 20 + 14, 1) = "*" + TM + "J" + "*"            '''''条码
        Excelapp.ActiveSheet.Cells(L * 20 + 18, 1) = "*" + TM + "J" + "*"            '''''条码
        Excelapp.ActiveSheet.Cells(L * 20 + 11, 2) = DT1.Recordset.Fields(10)                         ''''''筒颈
        Excelapp.ActiveSheet.Cells(L * 20 + 11, 4) = DT1.Recordset.Fields(6)                         ''''''克重
        
        Excelapp.ActiveSheet.Cells(L * 20 + 12, 2) = DT1.Recordset.Fields(12)                         ''''''颜色
        Excelapp.ActiveSheet.Cells(L * 20 + 12, 4) = cj   ''''DT1.Recordset.Fields(9)                         ''''''机台
        
        Excelapp.ActiveSheet.Cells(L * 20 + 13, 2) = k                           ''''''匹号
        Excelapp.ActiveSheet.Cells(L * 20 + 13, 4) = ""                          ''''重量
        Excelapp.ActiveSheet.Cells(L * 20 + 15, 4) = jh                           ''''''机台编号
        
dt2.RecordSource = "SELECT 纱支,批次,产地 FROM sxpbf WHERE 织号='" & DH & "' and 状态='用'"
dt2.Refresh
      
If Not dt2.Recordset.EOF Then  ''''''''''''''''''''''''''''''''''''
dt2.Recordset.MoveFirst
m = 0
Do While Not dt2.Recordset.EOF
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 1) = dt2.Recordset.Fields(0)  '''棉纱
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 3) = dt2.Recordset.Fields(1)  '''批次
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 4) = dt2.Recordset.Fields(2)  '''产地
dt2.Recordset.MoveNext
m = m + 1
Loop
Else     '''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''
DT1.RecordSource = "SELECT 纱支,批次,产地 FROM sxpb WHERE 织号='" & DH & "'"
DT1.Refresh
      
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
m = 0
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 1) = DT1.Recordset.Fields(0)  '''棉纱
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 3) = DT1.Recordset.Fields(1)  '''批次
        Excelapp.ActiveSheet.Cells(L * 21 + 7 + m, 4) = DT1.Recordset.Fields(2)  '''产地
DT1.Recordset.MoveNext
m = m + 1
Loop
End If
End If  ''''''''''''''''''''''''
End If

Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = False
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintOut
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub
Ert:

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Public Sub jhlcd(DT1 As Adodc, dt2 As Adodc, Zh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴针织软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\纺织\广兴\生产卡.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,款号,织号,品名,筒颈,计划,克重,幅宽,颜色,开幅线,备注,交期,日期,序号,纱别,车间,机台,匹重 FROM v_kpd_ctjh where 单据='" & Zh & "' order by 序号"
DT1.Refresh

If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit

Else
DT1.Recordset.MoveFirst
i = 2
L = 2
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(16)  ''''机号
Excelapp.ActiveSheet.Cells(2, 4) = DT1.Recordset.Fields(4)  ''''规格
Excelapp.ActiveSheet.Cells(2, 6) = Trim(DT1.Recordset.Fields(2))  ''''织号

Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(3)   '''品名
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(0)   '''客户
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(5)     '''''''''计划

Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)         '''匹重
Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(9)   ''''''开幅线
'Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(1)                        ''''款号



i = i + 1
DT1.Recordset.MoveNext
Loop
End If


DT1.RecordSource = "SELECT 织号,sum(计划) FROM v_kpd_ctjh where 单据='" & Zh & "' group by 织号 order by 织号"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
dt2.RecordSource = "SELECT 织号,纱支,批次,织耗,配比,备注,颜色,产地 FROM sxpb where 织号='" & DT1.Recordset.Fields(0) & "' order by 序号"
dt2.Refresh
If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(4, L) = dt2.Recordset.Fields(1)         '''纱支
Excelapp.ActiveSheet.Cells(5, L) = dt2.Recordset.Fields(7)                        ''''批次
Excelapp.ActiveSheet.Cells(7, L) = dt2.Recordset.Fields(2)                        ''''产地
L = L + 1
dt2.Recordset.MoveNext
Loop
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


