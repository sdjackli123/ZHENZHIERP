Attribute VB_Name = "车间生产"

Public Sub dmd(DT1 As Adodc, dt2 As Adodc, DH As String, DD As String)   ''''无标题

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

DT1.RecordSource = "SELECT 客户全称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区  FROM v_bmd WHERE 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  order by 匹号"
DT1.Refresh
If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 7) = Trim(DT1.Recordset.Fields(10)) '日期
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0) ''客户全称
        Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(3) ''品名
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2) ''款号
        Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(1) ''锅号
        Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(8) ''班次
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(5) ''色别
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4) ''光坯幅宽
        Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(7) ''克重
'i = 1
'L = 0
'Do While Not DT1.Recordset.EOF
'If Int(i / 19) = i / 19 And i > 0 Then
'i = 1
'L = L + 1
'End If
'       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
'       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
'       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
'i = i + 1
'DT1.Recordset.MoveNext

'Loop

'DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from v_bmd where 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  group by 重量"
'DT1.Refresh
dt2.RecordSource = "select * from ckgl where 单据号='" & DH & "' "     ''这里单据号必须等于DH才能调出数据"
dt2.Refresh
If Not dt2.Recordset.EOF Then


'mpzl = 0
'mpps = 0
'gpzl = 0
'gpms = 0
'If Not IsNull(DT1.Recordset.Fields(0)) Then
'DT1.Recordset.MoveFirst
'Do While Not DT1.Recordset.EOF
'mpzl = mpzl + Val(DT1.Recordset.Fields(0))
'mpps = mpps + Val(DT1.Recordset.Fields(1))
'gpzl = gpzl + Val(DT1.Recordset.Fields(2))
'gpms = gpms + Val(DT1.Recordset.Fields(3))
'DT1.Recordset.MoveNext
'Loop
'End If
Excelapp.ActiveSheet.Cells(14, 1) = dt2.Recordset.Fields(15) ''幅宽明细
Excelapp.ActiveSheet.Cells(32, 3) = dt2.Recordset.Fields(3) ''毛坯重量
Excelapp.ActiveSheet.Cells(10, 6) = dt2.Recordset.Fields(4) ''毛坯匹数
Excelapp.ActiveSheet.Cells(32, 6) = dt2.Recordset.Fields(3) ''毛坯重量
Excelapp.ActiveSheet.Cells(10, 8) = dt2.Recordset.Fields(12) ''来料单位
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
Public Sub dmdms(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\mdms.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区,码数  FROM bmd WHERE 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  order by 匹号"
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
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from bmd where 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(32, 3) = gpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps
Excelapp.ActiveSheet.Cells(32, 5) = gpms
Excelapp.ActiveSheet.Cells(32, 7) = mpzl


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
Public Sub dmd100(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户全称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区  FROM v_bmd WHERE 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  order by 匹号"
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
If Int(i / 34) = i / 34 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Val(DT1.Recordset.Fields(6))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from v_bmd where 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 3) = gpzl
'Excelapp.ActiveSheet.Cells(47, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps



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
Public Sub dmd100ms(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100ms.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区,码数  FROM bmd WHERE 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  order by 匹号"
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
If Int(i / 34) = i / 34 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 2) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
       Excelapp.ActiveSheet.Cells(i + 13, L * 3 + 3) = DT1.Recordset.Fields(11)
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from bmd where 锅号='" & DH & "' and 定型='大定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 3) = gpzl
Excelapp.ActiveSheet.Cells(47, 5) = gpms
'Excelapp.ActiveSheet.Cells(47, 7) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps


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

Public Sub xmdms(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\mdms.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区,码数  FROM bmd WHERE 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "'  order by 匹号"
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
If Int(i / 19) = i / 19 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from bmd where 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(32, 3) = gpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps
Excelapp.ActiveSheet.Cells(32, 5) = gpms
'Excelapp.ActiveSheet.Cells(32, 7) = mpzl


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
Public Sub xmd(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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

DT1.RecordSource = "SELECT 客户全称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区,码数  FROM v_bmd WHERE 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "' order by 匹号"
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
If Int(i / 19) = i / 19 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Trim(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from v_bmd where 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(32, 3) = gpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps
''Excelapp.ActiveSheet.Cells(32, 5) = gpms
'Excelapp.ActiveSheet.Cells(32, 6) = mpzl


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


Public Sub xmd100ms(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100ms.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期,布区,码数  FROM bmd WHERE 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "' order by 匹号"
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
If Int(i / 34) = i / 34 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from bmd where 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 3) = gpzl
Excelapp.ActiveSheet.Cells(47, 5) = gpms
'Excelapp.ActiveSheet.Cells(47, 7) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps


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
Public Sub xmd100(DT1 As Adodc, DH As String, DD As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户全称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期  FROM v_bmd WHERE 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "' order by 匹号"
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
If Int(i / 34) = i / 34 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Val(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select 重量,count(distinct 匹号),sum(光胚重量),sum(码数) from v_bmd where 锅号='" & DH & "' and 定型='小定型' and 缸号='" & DD & "'  group by 重量"
DT1.Refresh

mpzl = 0
mpps = 0
gpzl = 0
gpms = 0
If Not IsNull(DT1.Recordset.Fields(0)) Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(0))
mpps = mpps + Val(DT1.Recordset.Fields(1))
gpzl = gpzl + Val(DT1.Recordset.Fields(2))
gpms = gpms + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 3) = gpzl
''Excelapp.ActiveSheet.Cells(47, 5) = gpms
'Excelapp.ActiveSheet.Cells(47, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps


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


Public Sub dbq(DT1 As Adodc, DH As String, ph As Integer, DD As String, fs As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\bq.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT *  FROM v_bmd WHERE 锅号='" & DH & "' and 匹号='" & ph & "'  and 缸号='" & DD & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        If IsNull(DT1.Recordset.Fields(18)) Then
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''客户简称
        Else
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(18) ''客户全称
        End If
        Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(12) ''''匹号
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1)  ''''锅号
        ''Excelapp.ActiveSheet.Cells(5, 5) = Format(DT1.Recordset.Fields(9), "#0.0") ''''重量
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(8)    '''''色别
        Excelapp.ActiveSheet.Cells(6, 5) = Format(DT1.Recordset.Fields(16), "#0.0")   '''''''码数
        ''Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)    '''''款号
        Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(5)    '''''幅宽
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(3)    '''''品名
        Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(10)    '''''克重
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(27)    '''''班次
        Excelapp.ActiveSheet.Cells(5, 5) = Trim(DT1.Recordset.Fields(13))    '''''时间
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(33) '''员工（操作）
Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintOut Copies:=fs   '''''打印份数
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing


End Sub
Public Sub dbqxs(DT1 As Adodc, DH As String, ph As Integer, DD As String, xh As Integer)   ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\bqxs.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT *  FROM v_bmd WHERE 锅号='" & DH & "' and 匹号='" & ph & "' and 定型='大定型' and 单号='" & DD & "' and 序号='" & xh & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        Excelapp.ActiveSheet.Cells(1, 5) = DT1.Recordset.Fields(19)
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(21)
        Excelapp.ActiveSheet.Cells(4, 5) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(16)    '''''码数
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(5)
        
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(8, 5) = DT1.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(9, 2) = DT1.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(9, 5) = Format(DT1.Recordset.Fields(13), "mm-dd")


Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintOut
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Excelapp.Quit
Set Excelapp = Nothing


End Sub


Public Sub xbq(DT1 As Adodc, DH As String, DD As String, fs As Integer, fk As String)
    Dim i As Integer
    On Error GoTo Ert

    Dim Excelapp As Excel.Application
    Set Excelapp = New Excel.Application
    Excelapp.Visible = True ' 令Excel可见
    Excelapp.DisplayAlerts = False ' 关闭Excel警告消息

    ' 打开模板工作簿
    Dim wb As Excel.Workbook
    Set wb = Excelapp.Workbooks.Open(App.Path & "\打印模版\广兴\bq.xls")
    Dim ws As Excel.Worksheet
    Set ws = wb.Sheets(1)

    DT1.RecordSource = "SELECT * FROM v_bmd WHERE 锅号='" & DH & "' AND 缸号='" & DD & "'"
    DT1.Refresh
    DT1.Recordset.MoveFirst

    ' 更新匹号并逐个打印
    For i = 1 To fs
        With ws
            If IsNull(DT1.Recordset.Fields(18)) Then
                .Cells(3, 2) = DT1.Recordset.Fields(0)
            Else
                .Cells(3, 2) = DT1.Recordset.Fields(18)
            End If

            .Cells(7, 5) = i  '''' 匹号，从1到fs
            .Cells(4, 2) = DT1.Recordset.Fields(1) '''' 锅号
            .Cells(5, 2) = DT1.Recordset.Fields(8) ''''' 色别
            .Cells(3, 5) = fk       ''''' 幅宽
            .Cells(6, 2) = DT1.Recordset.Fields(3) ''''' 品名
            .Cells(4, 5) = DT1.Recordset.Fields(10) ''''' 克重
            .Cells(7, 2) = DT1.Recordset.Fields(27) ''''' 班次
            .Cells(5, 5) = Trim(DT1.Recordset.Fields(13)) ''''' 时间
        End With
        ws.PrintOut Copies:=1, Collate:=True ' 打印当前工作表，1份
    Next i
    
    ' 清理和退出
    Excelapp.Quit
    Set Excelapp = Nothing
    Set wb = Nothing
    Exit Sub

Ert:
    ' 错误处理，确保Excel应用正确关闭
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
    If Not wb Is Nothing Then
        Set wb = Nothing
    End If
End Sub





Public Sub CLBB(Flex As VSFlexGrid, fd1, BT As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\bbdy.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q = 0
        Q1 = 0

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
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         End If
         Next i
         End With

Excelapp.ActiveSheet.Cells(1, 1) = BT
Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1

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

Public Sub CLDY(Flex As VSFlexGrid, BT As String, Flex1 As VSFlexGrid)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\RSDY.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q = 0
        Q1 = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows



          For i = 1 To k + 1

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
         Next i
         End With

x = i + 1

        With Flex1

                n = .Rows


          For i = 1 To n + 1

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(x + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
         x = x + 1
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

Public Sub gzdc(Flex As VSFlexGrid, FD, BT As String)   ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\bbdy.xls")
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

Excelapp.ActiveSheet.Cells(i, 1) = "合计金额"
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



Public Sub dmd100dc(DT1 As Adodc, DH As String, pm As String)   ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期  FROM bmd WHERE 锅号='" & DH & "' and 定型='大定型' and 品名='" & pm & "' order by 匹号"
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
If Int(i / 34) = i / 34 And i > 0 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Val(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select distinct 品名,光胚幅宽,重量,匹数 from bmd where 锅号='" & DH & "' and 定型='大定型' and 品名='" & pm & "'"
DT1.Refresh

mpzl = 0
mpps = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(2))
mpps = mpps + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps

DT1.RecordSource = "select distinct 备注 from kpd where 锅号='" & DH & "' and 品名='" & pm & "'"
DT1.Refresh

mpbz = ""
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpbz = mpbz + DT1.Recordset.Fields(0)
DT1.Recordset.MoveNext
Loop
End If
Excelapp.ActiveSheet.Cells(4, 2) = mpbz

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

Public Sub xmd100dc(DT1 As Adodc, DH As String, pm As String)   ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\md100.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,款号,品名,光胚幅宽,色别,光胚重量,克重,班次,匹号,日期  FROM bmd WHERE 锅号='" & DH & "' and 定型='小定型' and 品名='" & pm & "' order by 匹号"
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
If Int(i / 34) = i / 34 Then
i = 1
L = L + 1
End If
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 1) = DT1.Recordset.Fields(9)
       Excelapp.ActiveSheet.Cells(i + 13, L * 2 + 2) = Val(DT1.Recordset.Fields(6))
i = i + 1
DT1.Recordset.MoveNext

Loop

DT1.RecordSource = "select distinct 品名,光胚幅宽,重量,匹数 from bmd where 锅号='" & DH & "' and 定型='小定型' and 品名='" & pm & "'"
DT1.Refresh

mpzl = 0
mpps = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpzl = mpzl + Val(DT1.Recordset.Fields(2))
mpps = mpps + Val(DT1.Recordset.Fields(3))
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveSheet.Cells(47, 6) = mpzl
Excelapp.ActiveSheet.Cells(10, 6) = mpps

DT1.RecordSource = "select distinct 备注 from kpd where 锅号='" & DH & "' and 品名='" & pm & "'"
DT1.Refresh

mpbz = ""
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
mpbz = mpbz + DT1.Recordset.Fields(0)
DT1.Recordset.MoveNext
Loop
End If
Excelapp.ActiveSheet.Cells(4, 2) = mpbz


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

Public Sub mpbq(DT1 As Adodc, DH As String, xh As Integer)     ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\mpbq.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT 客户名称,锅号,色别,标签,光胚幅宽,品名,技术要求  FROM kpd WHERE 锅号='" & DH & "' and ip='" & xh & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(0)  '''客户
'        Excelapp.ActiveSheet.Cells(4, 5) = dt1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)  ''''缸号
        Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(2)  ''''缸号
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(3)  ''''款号
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(4)  ''''幅宽
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(5)  ''''品名
        Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(6)  ''''克重
        

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
Public Sub mbdy(DT1 As Adodc, selectedGuoHao As String, sh As Excel.Worksheet)
    On Error GoTo errorhandler

    ' 确认 DT1 已正确连接到数据库
    If DT1.ConnectionString = "" Then
        MsgBox "DT1 is not connected to the database."
        Exit Sub
    End If

    ' 查询数据库获取客户、日期、色别和色名
    Debug.Print "Executing SQL query through DT1" ' 打印调试信息
    DT1.RecordSource = "SELECT 客户名称, 日期, 色别, 色名 FROM v_kpd_khmb WHERE 锅号='" & selectedGuoHao & "'"
    DT1.Refresh

    ' 确保记录集存在
    If DT1.Recordset Is Nothing Then
        MsgBox "SQL query failed. Check the database connection and query."
        GoTo Cleanup
    End If

    ' 初始化 rowIndex
    Dim rowIndex As Integer
    rowIndex = sh.Cells(sh.Rows.count, 1).End(xlUp).Row + 1
    Debug.Print "Row index initialized to: " & rowIndex ' 打印调试信息

    If Not DT1.Recordset.EOF Then
        Debug.Print "Record found for GuoHao: " & selectedGuoHao ' 打印调试信息
        sh.Cells(rowIndex, 1).value = DT1.Recordset.Fields("客户名称").value
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "日期：" & DT1.Recordset.Fields("日期").value & " 到货明细"
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "颜色：" & DT1.Recordset.Fields("色别").value & " 色号：" & DT1.Recordset.Fields("色名").value
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "如有信息修正，请及时告知，请确认色号"
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "" ' 空一行，确保信息之间有空行
        rowIndex = rowIndex + 1
    Else
        Debug.Print "No records found for GuoHao: " & selectedGuoHao ' 打印调试信息
    End If

Cleanup:
    Exit Sub

errorhandler:
    MsgBox "错误: " & Err.Description
    Resume Cleanup
End Sub

