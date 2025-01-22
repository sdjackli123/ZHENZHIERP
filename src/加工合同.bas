Attribute VB_Name = "加工合同"
Public Sub htht(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")


Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\打印模版合同.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_cpjf where 订单编号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(4, 2) = gh
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(6, 2) = Trim(DT1.Recordset.Fields(5))
Excelapp.ActiveSheet.Cells(85, 4) = DT1.Recordset.Fields(1)  ''''''交货地址
Excelapp.ActiveSheet.Cells(86, 4) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(87, 4) = DT1.Recordset.Fields(3)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct 客户,日期,交期  from sczy_x where 单号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(66, 3) = CDate(DT1.Recordset.Fields(2)) - CDate(DT1.Recordset.Fields(1))
Excelapp.ActiveSheet.Cells(66, 5) = Trim(DT1.Recordset.Fields(1))
Excelapp.ActiveSheet.Cells(66, 8) = Trim(DT1.Recordset.Fields(2))
End If


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_mlgg where 订单编号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(15, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(15, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(15, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(15, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(15, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(15, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(15, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(15, 10) = DT1.Recordset.Fields(9)
End If


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_flgg where 订单编号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(16, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(16, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(16, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(16, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(16, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(16, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(16, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(16, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(16, 10) = DT1.Recordset.Fields(9)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select *  from htfz_qtgg where 订单编号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(17, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(17, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(17, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(17, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(17, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(17, 7) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(17, 8) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(17, 9) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(17, 10) = DT1.Recordset.Fields(9)
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct 色别,色牢度  from sczy_x where 单号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 26
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(2)
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select distinct 品名,幅宽,克重,缩水率,扭度,布纹  from sczy_x where 单号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 40
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(5)
i = i + 1
DT1.Recordset.MoveNext
Loop
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from htfz_cpbmyq where 订单编号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(52, 1) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(52, 5) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(52, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(58, 1) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(120, 2) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(145, 4) = DT1.Recordset.Fields(6)
End If

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from htfz_qybyj where 订单编号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(63, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(64, 2) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(82, 3) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(83, 3) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(92, 3) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(108, 3) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(109, 3) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(110, 3) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(110, 7) = DT1.Recordset.Fields(9)
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


Public Sub DXDY(DT1 As Adodc, dt3 As Adodc, DH As String, xh1 As Integer, xh2 As Integer, ps As Integer, sl As Single) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\委外出库.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,品名,色别,锅号,'' as 和约号,匹数,重量,'' as 单价,'' as 金额,日期,备注,类别,光坯  FROM wwkpd WHERE 单据='" & DH & "' and 序号 between '" & xh1 & "' and '" & xh2 & "' order BY 序号"
DT1.Refresh


DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   '''客户名称
        Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(9))   '''日期
        Excelapp.ActiveSheet.Cells(2, 12) = DH    '''单号
        Excelapp.ActiveSheet.Cells(9, 6) = ps   '''匹数
        Excelapp.ActiveSheet.Cells(9, 7) = sl   '''数量

i = 4
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(2)    ''''色别
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(3)    ''''锅号
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(5)    ''''匹数
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(6)    ''''数量
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(12)    ''''光坯
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(11)    '''''加工类别
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(10)    '''''备注
i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "SELECT 模块 FROM yhb WHERE 用户='" & yhm & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
        Excelapp.ActiveSheet.Cells(11, 12) = DT1.Recordset.Fields(0)   '''制单
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

Public Sub XSHT(DT1 As Adodc, DH As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\销售合同.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,单号,款号,品名,幅宽+克重,色别,计划,单价,备注,日期,交期,序号,总备注,投染类别,面料用途  FROM sczykpd WHERE 单号='" & DH & "' order BY 序号"
DT1.Refresh


If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(0)   '''客户名称
        Excelapp.ActiveSheet.Cells(6, 7) = DT1.Recordset.Fields(1)   '''合同编号
        Excelapp.ActiveSheet.Cells(7, 7) = Trim(DT1.Recordset.Fields(9))    '''合同日期

i = 10
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    ''''品名
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(4)    ''''规格
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(5)    ''''颜色
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)    ''''数量
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(7)    ''''单价
        Excelapp.ActiveSheet.Cells(i, 6) = Format(DT1.Recordset.Fields(6) * DT1.Recordset.Fields(7), "#0.00") ''''金额
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(8)    '''''备注
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

Public Sub SCTZD(DT1 As Adodc, DH As String)  ''''无标题

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert
        Dim L As Integer
        Dim ym As Integer
        Dim dym As Integer

        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

       On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\生产通知单.xls")
'5)设置第2个工作表为活动工作表：

DT1.RecordSource = "SELECT 序号 FROM sczykpd WHERE 单号='" & DH & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then

L = DT1.Recordset.RecordCount
If L / 10 <> Int(L / 10) Then
ym = Int(L / 10) + 1
Else
ym = Int(L / 10)
End If
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户,单号,款号,品名,幅宽,克重,色别,计划,色名,备注,日期,交期,序号,总备注,投染类别,面料用途,成分,负责,isnull(意见,'') as 确认意见  FROM sczykpd WHERE 单号='" & DH & "' order BY 序号"
DT1.Refresh

dym = 1
L = 1
DT1.Recordset.MoveFirst
        
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(ym)   '''页数
        Excelapp.ActiveSheet.Cells(3, 16) = Trim(dym)   '''第几页
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   '''客户名称
        Excelapp.ActiveSheet.Cells(4, 13) = Trim(DT1.Recordset.Fields(10))   '''合同日期
        Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(1)    '''合同编号
        Excelapp.ActiveSheet.Cells(5, 13) = DT1.Recordset.Fields(17)    ''合同业务
        Excelapp.ActiveSheet.Cells(20, 15) = DT1.Recordset.Fields(15)    '''面料用途
        Excelapp.ActiveSheet.Cells(24, 15) = Trim(DT1.Recordset.Fields(11))    '''交期
        Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(13)    '''总备注

i = 7
yj = ""
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)    '''品名
        Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(16)    '''成分
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(8)    ''''色号
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)    ''''颜色
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)    ''''门幅
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(5)    ''''克重
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(7)    ''''数量
        Excelapp.ActiveSheet.Cells(i, 13) = DT1.Recordset.Fields(9)  ''''备注
yj = Trim(yj) + Trim(i) + DT1.Recordset.Fields(18)
i = i + 1
L = L + 1
DT1.Recordset.MoveNext
If L = 11 Then
        Excelapp.ActiveSheet.Cells(19, 1) = yj    '''确认意见
        yj = ""
dym = dym + 1
L = 1
i = 7
        Excelapp.Sheets(dym).Activate
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(ym)   '''页数
        Excelapp.ActiveSheet.Cells(3, 16) = Trim(dym)   '''第几页
        Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   '''客户名称
        Excelapp.ActiveSheet.Cells(4, 13) = Trim(DT1.Recordset.Fields(10))   '''合同日期
        Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(1)    '''合同编号
        Excelapp.ActiveSheet.Cells(5, 13) = DT1.Recordset.Fields(17)    ''合同业务
        Excelapp.ActiveSheet.Cells(20, 15) = DT1.Recordset.Fields(15)    '''面料用途
        Excelapp.ActiveSheet.Cells(24, 15) = Trim(DT1.Recordset.Fields(11))    '''交期
        Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(13)    '''总备注
End If
Loop
        Excelapp.ActiveSheet.Cells(19, 1) = yj    '''确认意见
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



