Attribute VB_Name = "毛坯仓库"
Public Password As String
Public riqi As String
Public color As String
Public guohao As String
Public cunt As Single
Public sehao As String
Public shjian As String
Public fweishu As Integer
Public fzh(40) As String
Public bsh(40) As Integer
Public passwordzhu As String
Public user As String
Public passbiao As Integer
Public ndr As String
Public cpf As String
Public BDRQ As String  ''''''日期变量流程卡录入

Public pu As Integer
Public zu(10) As String   ''''''''''''磅码用交换变量
Public pd As Integer  ''''adodc6jilushu
Public khmc As String
Public zhlhji As Long
Public ww As Integer '操作是否进行
Public DH As String   '单号变量，用于复制
Public Sub fhbb1(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, BT As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\fhbb.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
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
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         Q6 = Val(Excelapp.ActiveSheet.Cells(i, fd6)) + Q6
         Q7 = Val(Excelapp.ActiveSheet.Cells(i, fd7)) + Q7
         Q8 = Val(Excelapp.ActiveSheet.Cells(i, fd8)) + Q8
        End If
         Next i


        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计数量"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5
Excelapp.ActiveSheet.Cells(i, fd6) = Q6
Excelapp.ActiveSheet.Cells(i, fd7) = Q7
Excelapp.ActiveSheet.Cells(i, fd8) = Q8
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
Public Sub mprk(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\毛坯入库.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select 客户名称,布类,和约号,毛胚幅宽,毛胚匹数,毛胚重量,备注,负责人,日期,ny,颜色,克重,业务,备注,幅宽明细 from ckgl where 单据号='" & gh & "'  order by ip"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) ''客户
Excelapp.ActiveSheet.Cells(3, 5) = Trim(DT1.Recordset.Fields(8)) ''日期
Excelapp.ActiveSheet.Cells(3, 8) = Trim(gh) ''单据号
Excelapp.ActiveSheet.Cells(3, 11) = DT1.Recordset.Fields(7) ''负责
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(12) ''司机
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1) '''布类
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(10)    ''''''''''颜色
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(11)    ''''''''''克重
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(13)    ''''''''''备注

' 引用必要的正则表达式库
Dim regexWidth As Object, regexNumber As Object
Set regexWidth = CreateObject("VBScript.RegExp")
regexWidth.Global = True
regexWidth.Pattern = "(\b\d{1,2}cm\b|领|袖)" ' 匹配数字后跟"cm"或“领”或“袖”

Set regexNumber = CreateObject("VBScript.RegExp")
regexNumber.Global = True
regexNumber.Pattern = "\d+(\.\d+)?" ' 匹配数字，包括小数

' 读取幅宽明细
Dim details As String
details = DT1.Recordset.Fields(14) ' 假设幅宽明细在第15个字段

' 找到所有的幅宽标签及其对应重量
Dim widths As Object, weights As Object, pieces As Object, totalWeights As Object
Set widths = CreateObject("Scripting.Dictionary")
Set weights = CreateObject("Scripting.Dictionary")
Set pieces = CreateObject("Scripting.Dictionary")
Set totalWeights = CreateObject("Scripting.Dictionary")

Dim currentWidth As String, currentWeights As String, pieceCount As Integer, weightSum As Double
currentWidth = ""
currentWeights = ""
pieceCount = 0
weightSum = 0
Dim detailArray() As String
detailArray = Split(details, " ") ' 使用空格分割各部分

' 遍历分割后的数据
For i = LBound(detailArray) To UBound(detailArray)
    If regexWidth.Test(detailArray(i)) Then
        If currentWidth <> "" Then
            ' 存储上一个幅宽的匹数、重量和明细
            pieces.Add currentWidth, pieceCount
            totalWeights.Add currentWidth, weightSum
            weights.Add currentWidth, Trim(currentWeights)
            ' 重置计数和重量累计
            pieceCount = 0
            weightSum = 0
            currentWeights = ""
        End If
        currentWidth = detailArray(i)
        ' 存储幅宽
        widths.Add currentWidth, currentWidth
    ElseIf regexNumber.Test(detailArray(i)) Then
        ' 累计当前幅宽的匹数和重量
        pieceCount = pieceCount + 1
        weightSum = weightSum + CDbl(detailArray(i))
        currentWeights = currentWeights & detailArray(i) & " "
    End If
Next i
' 存储最后一个幅宽的匹数、重量和明细
If currentWidth <> "" Then
    pieces.Add currentWidth, pieceCount
    totalWeights.Add currentWidth, weightSum
    weights.Add currentWidth, Trim(currentWeights)
End If

' 输出幅宽到Excel单元格
Dim col As Integer
col = 2 ' 从第2列开始
For Each Key In widths.Keys
    Excelapp.ActiveSheet.Cells(5, col).value = widths(Key)
    Excelapp.ActiveSheet.Cells(6, col).value = pieces(Key) ' 匹数
    Excelapp.ActiveSheet.Cells(7, col).value = totalWeights(Key) ' 总重量
    Excelapp.ActiveSheet.Cells(9, col).value = weights(Key) ' 重量明细
    col = col + 2 ' 假设每个幅宽占用两列，可以根据需要调整
Next Key

End If

DT1.RecordSource = "select sum(毛胚匹数),round(sum(毛胚重量),2) from ckgl where 单据号='" & gh & "'"
DT1.Refresh
'Excelapp.ActiveSheet.Cells(11, 5) = DT1.Recordset.Fields(0)
'Excelapp.ActiveSheet.Cells(11, 6) = DT1.Recordset.Fields(1)

Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
'MsgBox ("")
        Excelapp.DisplayAlerts = False
       ' Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub mprkf(DT1 As Adodc, gh As String, xh1 As Integer, xh2 As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\毛坯入库分.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select 客户名称,布类,和约号,毛胚幅宽,毛胚匹数,毛胚重量,备注,负责人,日期,ny from ckgl where 单据号='" & gh & "' and ip between '" & xh1 & "' and '" & xh2 & "' order by ip"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 5) = Trim(DT1.Recordset.Fields(8))
Excelapp.ActiveSheet.Cells(3, 8) = Trim(gh)
Excelapp.ActiveSheet.Cells(12, 2) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(12, 5) = DT1.Recordset.Fields(9)

i = 6
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ny来料单位
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(6)

i = i + 1
DT1.Recordset.MoveNext
Loop
End If

Excelapp.ActiveWindow.Zoom = 100
        Excelapp.Visible = True
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
Public Sub mpck(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "彩虹打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\mpck.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select 客户,布类,款号,毛胚幅宽,毛胚匹数,毛胚重量,备注,配缸负责,出库日期 from mpbh where 锅号='" & gh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 5) = Trim(DT1.Recordset.Fields(8))
Excelapp.ActiveSheet.Cells(3, 8) = Trim(gh)
Excelapp.ActiveSheet.Cells(13, 3) = DT1.Recordset.Fields(7)

i = 6
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(6)

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

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
Public Sub lcd3(DT1 As Adodc, dt2 As Adodc, gh As String, xh As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\毛坯码单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


dt2.RecordSource = "select 客户名称,单号,锅号,品名,光胚幅宽,技术要求,色别,匹数,重量,备注,标签 from kpd where 锅号='" & gh & "' and ip='" & xh & "'"
dt2.Refresh

If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(4, 2) = dt2.Recordset.Fields(9)    '''备注
Excelapp.ActiveSheet.Cells(6, 2) = dt2.Recordset.Fields(0)    '''客户
Excelapp.ActiveSheet.Cells(6, 6) = dt2.Recordset.Fields(3)    ''''品名
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(1)    ''''单号
Excelapp.ActiveSheet.Cells(8, 4) = dt2.Recordset.Fields(2)    ''''锅号
Excelapp.ActiveSheet.Cells(8, 8) = dt2.Recordset.Fields(6)    ''''色别
Excelapp.ActiveSheet.Cells(10, 2) = dt2.Recordset.Fields(4)    ''''幅宽
Excelapp.ActiveSheet.Cells(10, 4) = dt2.Recordset.Fields(5)    ''''克重
Excelapp.ActiveSheet.Cells(10, 6) = dt2.Recordset.Fields(7)    ''''匹数
Excelapp.ActiveSheet.Cells(10, 8) = dt2.Recordset.Fields(10)    ''''客户单号  就是款号

DT1.RecordSource = "select 交期,备注,技要,日期 from sczy_x where 单号='" & dt2.Recordset.Fields(1) & "' and 序号='" & xh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(6, 4) = Trim(DT1.Recordset.Fields(3))    '''计划日期
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1) + DT1.Recordset.Fields(2)   '''备注技要
End If

m = 0
k = 14  '''''''''''''''''''''''''

DT1.RecordSource = "select 匹号,重量 from mpbmd where 锅号='" & gh & "' and 序号='" & xh & "'"
DT1.Refresh

Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(k, 1 + m * 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(k, 2 + m * 2) = DT1.Recordset.Fields(1)
k = k + 1
If k = 32 Then
m = m + 1
k = 14
End If
DT1.Recordset.MoveNext
Loop
End If



'Excelapp.ActiveSheet.Cells(3, 9) = Mid(L, 1, Len(L) - 1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Excelapp.ActiveSheet.Cells(3 + 22, 9) = Mid(L, 1, Len(L) - 1)


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





Public Sub lcd(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\LCdD.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = cast('" & a & "' as real)  "
DT1.Refresh
Excelapp.ActiveSheet.Cells(2, 6) = DT1.Recordset.Fields(12)
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 12) = DT1.Recordset.Fields(8)
DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' ORDER BY IP "
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
End If
DT1.Recordset.MoveFirst
i = 5
Do While Not DT1.Recordset.EOF

Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 2) = DT1.Recordset.Fields(4) + "/" + DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(6)
Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(7)
Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(9)

DT1.Recordset.MoveNext
i = i + 1
Loop



Excelapp.ActiveWindow.Zoom = 200

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

Public Sub TMDY(TM As String, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\TMDY.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate
Excelapp.ActiveSheet.Cells(1, 1) = TM
Excelapp.ActiveSheet.Cells(4, 1) = gh
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

Public Sub ZJTMDY(TM As String, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\TMDY.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate
Excelapp.ActiveSheet.Cells(1, 1) = TM
Excelapp.ActiveSheet.Cells(4, 1) = gh
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

Public Sub lcd2(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")
Dim lc As String


Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\排缸卡ok.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & a & "'"
DT1.Refresh
Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   ''''客户
Excelapp.ActiveSheet.Cells(2, 8) = "*" + DT1.Recordset.Fields(2) + "J*" '''条码
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(2)   '''锅号

Excelapp.ActiveSheet.Cells(6, 3) = Trim(DT1.Recordset.Fields(12))    ''''日期
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(6, 7) = "返修"
Else
Excelapp.ActiveSheet.Cells(6, 7) = "正常"
End If '''' 类别
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''款号
Excelapp.ActiveSheet.Cells(10, 3) = DT1.Recordset.Fields(52)     ''色号
Excelapp.ActiveSheet.Cells(10, 7) = DT1.Recordset.Fields(8)     '''''颜色

Excelapp.ActiveSheet.Cells(14, 3) = DT1.Recordset.Fields(14)     ''机台
   

''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(重量,0)),2),SUM(isnull(匹数,0)) from kpd where 锅号='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(1)   ''''匹数
Excelapp.ActiveSheet.Cells(12, 7) = DT1.Recordset.Fields(0)    ''''计划量
End If

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(25 + i * 1, 2) = DT1.Recordset.Fields(55)   '''编号
Excelapp.ActiveSheet.Cells(25 + i * 1, 3) = DT1.Recordset.Fields(3)   '''品名
Excelapp.ActiveSheet.Cells(25 + i * 1, 5) = DT1.Recordset.Fields(5)   '''幅宽
Excelapp.ActiveSheet.Cells(25 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''克重
Excelapp.ActiveSheet.Cells(25 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''匹数
Excelapp.ActiveSheet.Cells(25 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''重量
Excelapp.ActiveSheet.Cells(25 + i * 1, 9) = DT1.Recordset.Fields(9)       ''''备注 技术要求
i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct 编号,mr from kpd where  锅号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
lc = ""
Do While Not DT1.Recordset.EOF
lc = lc + "编号：" + DT1.Recordset.Fields(0) + " 流程：" + DT1.Recordset.Fields(1) + "/"  ''''''''''流程
DT1.Recordset.MoveNext
Loop
End If
Excelapp.ActiveSheet.Cells(20, 3) = lc   ''''流程

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
Public Sub lcd222(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\新流程单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = cast('" & a & "' as real)"
DT1.Refresh
Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(0)    ''''客户名称
'Excelapp.ActiveSheet.Cells(4, 4) = Trim(dt1.Recordset.Fields(12))   '''''日期
Excelapp.ActiveSheet.Cells(5, 17) = DT1.Recordset.Fields(2)   '''''锅号
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(8)  ''''本厂色名
Excelapp.ActiveSheet.Cells(6, 10) = DT1.Recordset.Fields(52)   '''''色别 客户色别
Excelapp.ActiveSheet.Cells(7, 3) = DT1.Recordset.Fields(13)   ''''''标签  客户单号
'Excelapp.ActiveSheet.Cells(3, 2) = dt1.Recordset.Fields(3)    '''''品名
DH = DT1.Recordset.Fields(1)
xh = DT1.Recordset.Fields(11)
Excelapp.ActiveSheet.Cells(1, 16) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '锅号条码
Excelapp.ActiveSheet.Cells(13, 3) = DT1.Recordset.Fields(30)  '''加工说明
Excelapp.ActiveSheet.Cells(15, 3) = DT1.Recordset.Fields(51)  '''加工要求



''''''''''''''''''''''''''''
DT1.RecordSource = "select * from sczy_z where 单号='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(18, 3) = DT1.Recordset.Fields(1)                ''''''''''总备注
Excelapp.ActiveSheet.Cells(6, 3) = DT1.Recordset.Fields(0)                ''''''''''本厂单号
End If

DT1.RecordSource = "select isnull(流程,'') from sczy_x where 单号='" & DH & "' and 序号='" & xh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(17, 3) = DT1.Recordset.Fields(0)                ''''''''''流程
End If


DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 9
Do While Not DT1.Recordset.EOF
dt2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
dt2.RecordSource = "select 备注,交期,日期 from sczy_x where 单号='" & DT1.Recordset.Fields(1) & "' and 序号='" & DT1.Recordset.Fields(11) & "'"
dt2.Refresh
If Not dt2.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(0)   '''织厂
Excelapp.ActiveSheet.Cells(4, 10) = Trim(dt2.Recordset.Fields(1))   '''交期
Excelapp.ActiveSheet.Cells(4, 4) = Trim(dt2.Recordset.Fields(2))   '''''日期
Else
Excelapp.ActiveSheet.Cells(i, 1) = ""   '''织厂
End If
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)   '''品名
Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)   '''匹数
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(7)   '''重量
Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)  '''克重
Excelapp.ActiveSheet.Cells(i, 18) = DT1.Recordset.Fields(5)  '''门幅

i = i + 1
DT1.Recordset.MoveNext
Loop
End If



DT1.RecordSource = "select round(SUM(重量),2),round(SUM(匹数),1) from kpd where 锅号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(6, 17) = DT1.Recordset.Fields(1)   '''合计匹数
Excelapp.ActiveSheet.Cells(7, 17) = DT1.Recordset.Fields(0)  ''合计重量
End If



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

Public Sub jdzt(Flex As VSFlexGrid, BT As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\jdzt.xls")
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




Public Sub jdmx(Flex As VSFlexGrid, BT As String)    ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\jdmx.xls")
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


Public Sub BBDY(Flex As VSFlexGrid, fd1, fd2, BT As String)  ''''无标题

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
        Q2 = 0

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
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         End If
         Next i
         End With

Excelapp.ActiveSheet.Cells(1, 1) = BT
Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2

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


Public Sub fhbb(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, BT As String) ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\fhbb.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
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
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         Q6 = Val(Excelapp.ActiveSheet.Cells(i, fd6)) + Q6
        End If
         Next i


        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计数量"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5
Excelapp.ActiveSheet.Cells(i, fd6) = Q6

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



Public Sub bhmx(Flex As VSFlexGrid, fd1, fd2, BT)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\bhmx.xls")
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

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 3 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2

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




Public Sub sx(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''调整显示格式
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
If Int(Val(MSFlex.Text)) = Val(MSFlex.Text) Then
MSFlex.Text = Int(Val(MSFlex.Text))
Else
MSFlex.Text = Int(Val(MSFlex.Text)) + 1
End If
Next
End Sub


Public Sub SX1(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''调整显示格式
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.0")
Next
End Sub

Public Sub SX2(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''调整显示格式
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.00")
Next
End Sub
Public Sub SX3(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''调整显示格式
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.000")
Next
End Sub
Public Sub SX4(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''调整显示格式
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.0000")
Next
End Sub




Public Sub PCOutadodcToExcel(Flex As VSFlexGrid)

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


       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
       L = 0
       m = 0
       n = 0
       Q = 0
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 2 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, 4)) + Q
         L = Val(Excelapp.ActiveSheet.Cells(i, 6)) + L
         m = Val(Excelapp.ActiveSheet.Cells(i, 8)) + m
         n = Val(Excelapp.ActiveSheet.Cells(i, 10)) + n
         End If
         Next i

        End With
'9) 在第8行之前插入分页符：



Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, 4) = Q
Excelapp.ActiveSheet.Cells(i, 6) = L
Excelapp.ActiveSheet.Cells(i, 8) = m
Excelapp.ActiveSheet.Cells(i, 10) = n


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
Public Sub MXOutadodcToExcel(Flex As VSFlexGrid, BT As String)

    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    Dim Q   As Double ' 使用Double类型来确保合计值可以处理小数
    On Error GoTo Ert

    Dim Excelapp   As Excel.Application

    ' 创建Excel应用程序实例
    Set Excelapp = New Excel.Application

    On Error Resume Next

    ' 设置新工作簿的工作表数量
    Excelapp.SheetsInNewWorkbook = 1
    
    ' 设置Excel应用程序标题
    Excelapp.Caption = "广兴打印模版软件之打印"
    
    ' 打开已存在的工作簿
    Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
    
    ' 激活第一个工作表
    Excelapp.Sheets(1).Activate
    
    Q = 0 ' 初始化合计值
    
    With Flex
        k = .Rows ' 获取表格的总行数

        ' 遍历表格中的数据
        For i = 1 To k
            For j = 1 To .Cols
                DoEvents
                
                ' 检查单元格的值是否为数字
                If IsNumeric(.TextMatrix(i - 1, j)) Then
                    ' 如果是数字，直接赋值为数字格式，去掉单引号
                    Excelapp.ActiveSheet.Cells(i + 1, j).value = CDbl(.TextMatrix(i - 1, j))
                Else
                    ' 如果不是数字，按文本格式处理
                    Excelapp.ActiveSheet.Cells(i + 1, j).value = .TextMatrix(i - 1, j)
                End If
            Next j
            
            ' 累加某列（FD列）的数值（确保FD列为有效列）
            If i >= 1 Then
                Q = Q + Val(.TextMatrix(i - 1, FD))
            End If
        Next i
    End With

    ' 将合计值精确到两位小数后导出到最后一行
    Excelapp.ActiveSheet.Cells(k + 2, FD).value = Format(Q, "0.00")

    ' 设置Excel单元格的标题
    Excelapp.ActiveSheet.Cells(1, 1) = BT
    
    ' 设置窗口缩放比例
    Excelapp.ActiveWindow.Zoom = 100
    
    ' 显示Excel应用程序
    Excelapp.Visible = True
    
    ' 禁用警告提示
    Excelapp.DisplayAlerts = False
    
    ' 退出并清除Excel应用程序实例
    Set Excelapp = Nothing
    Excelapp.Quit
    Exit Sub

Ert:
    ' 错误处理，退出Excel
    Set Excelapp = Nothing
    Excelapp.Quit

End Sub
Public Sub OutadodcToExcel(Flex As VSFlexGrid, FD, BT)

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴染整软件之打印"
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


Public Sub lyldy(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\履约率.xls")
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

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 1 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
'Excelapp.ActiveSheet.Cells(1, 1) = BT
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
Public Sub hzdy(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\整理.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(7).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 1 To k

                          For j = 1 To .Cols

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 1 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
'Excelapp.ActiveSheet.Cells(1, 1) = BT
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
Public Sub yrdy(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\印染.xls")
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

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 1 Then
         Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
'Excelapp.ActiveSheet.Cells(1, 1) = BT
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

Public Sub YEBDOutadodcToExcel(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''无标题

        On Error GoTo Ert
        Dim i   As Integer
        Dim TT   As Single
        Dim k   As Integer
        Dim Q As Single
        Dim L   As Integer

        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\YEB.xls")
'5)设置第1个工作表为活动工作表：
Excelapp.Sheets(3).Activate
DT1.RecordSource = "SELECT 会计科目 FROM ZLCX  GROUP BY 会计科目"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT 会计科目,借贷方向 FROM ZLCX  GROUP BY 会计科目,借贷方向"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''移动到第一条记录
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "页"   '''''页
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''期间



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "页" '''''页
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''期间
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  会计科目='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "结") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "汇总" Then
Excelapp.ActiveSheet.Cells(k, 4) = dt2.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(k, 5) = dt2.Recordset.Fields(5)
End If
       
If InStr(dt2.Recordset.Fields(1), "结") > 0 And dt2.Recordset.Fields(9) = "3" Then
Excelapp.ActiveSheet.Cells(k, 7) = Val(dt2.Recordset.Fields(7))
End If
dt2.Recordset.MoveNext
Loop
k = k + 1
Q = Q + 1
DT1.Recordset.MoveNext
Loop

Excelapp.ActiveWindow.Zoom = 100
Excelapp.Visible = True
MsgBox ("")
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

Public Sub OutadodcToExcel3(Flex As VSFlexGrid, fd1, fd2, fd3, BT) ''''按一字段合计（含标题）

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴制帽软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
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
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
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
Public Sub OutadodcToExcel22(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, fd9, fd10, fd11, fd12, fd13, BT) ''''按一字段合计（含标题）

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

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 3 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         Q5 = Val(Excelapp.ActiveSheet.Cells(i, fd5)) + Q5
         Q6 = Val(Excelapp.ActiveSheet.Cells(i, fd6)) + Q6
         Q7 = Val(Excelapp.ActiveSheet.Cells(i, fd7)) + Q7
         Q8 = Val(Excelapp.ActiveSheet.Cells(i, fd8)) + Q8
         Q9 = Val(Excelapp.ActiveSheet.Cells(i, fd9)) + Q9
         Q10 = Val(Excelapp.ActiveSheet.Cells(i, fd10)) + Q10
         Q11 = Val(Excelapp.ActiveSheet.Cells(i, fd11)) + Q11
         Q12 = Val(Excelapp.ActiveSheet.Cells(i, fd12)) + Q12
         Q13 = Val(Excelapp.ActiveSheet.Cells(i, fd13)) + Q13
        
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4
Excelapp.ActiveSheet.Cells(i, fd5) = Q5
Excelapp.ActiveSheet.Cells(i, fd6) = Q6
Excelapp.ActiveSheet.Cells(i, fd7) = Q7
Excelapp.ActiveSheet.Cells(i, fd8) = Q8
Excelapp.ActiveSheet.Cells(i, fd9) = Q9
Excelapp.ActiveSheet.Cells(i, fd10) = Q10
Excelapp.ActiveSheet.Cells(i, fd11) = Q11
Excelapp.ActiveSheet.Cells(i, fd12) = Q12
Excelapp.ActiveSheet.Cells(i, fd13) = Q13

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
Public Sub OutadodcToExcel2(Flex As VSFlexGrid, fd1, fd2, BT) ''''按一字段合计（含标题）

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

                              Excelapp.ActiveSheet.Cells(i + 1, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 3 Then
         Q1 = Val(Excelapp.ActiveSheet.Cells(i, fd1)) + Q1
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         
        
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

'Excelapp.ActiveSheet.Cells(i, 1) = "合计"
'Excelapp.ActiveSheet.Cells(i, fd1) = Q1
'Excelapp.ActiveSheet.Cells(i, fd2) = Q2


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

Public Sub OutadodcToExcel4(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, BT) ''''按一字段合计（含标题）

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "广兴制帽软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\lbj.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        

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
         Q2 = Val(Excelapp.ActiveSheet.Cells(i, fd2)) + Q2
         Q3 = Val(Excelapp.ActiveSheet.Cells(i, fd3)) + Q3
         Q4 = Val(Excelapp.ActiveSheet.Cells(i, fd4)) + Q4
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "合计"
Excelapp.ActiveSheet.Cells(i, fd1) = Q1
Excelapp.ActiveSheet.Cells(i, fd2) = Q2
Excelapp.ActiveSheet.Cells(i, fd3) = Q3
Excelapp.ActiveSheet.Cells(i, fd4) = Q4

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



Public Sub lcd33(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\排缸卡.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select distinct 锅号,客户名称,CONVERT(varchar,日期, 23),标签 from kpd where 单号='" & gh & "' order by 锅号"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(2, 13) = Trim(gh)

i = 4
Do While Not DT1.Recordset.EOF

dt2.RecordSource = "select * from kpd where 锅号='" & DT1.Recordset.Fields(0) & "' order by IP"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(i, 4) = dt2.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(i, 5) = dt2.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(i, 6) = Trim(dt2.Recordset.Fields(6))
Excelapp.ActiveSheet.Cells(i, 7) = Trim(dt2.Recordset.Fields(7))
Excelapp.ActiveSheet.Cells(i, 8) = dt2.Recordset.Fields(14)
Excelapp.ActiveSheet.Cells(i, 9) = dt2.Recordset.Fields(9)
Excelapp.ActiveSheet.Cells(i, 13) = dt2.Recordset.Fields(10)
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

Public Sub lcd2222(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\流程单ok.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & a & "'"
DT1.Refresh
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(2, 4) = DT1.Recordset.Fields(15)
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(14)
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(12))
Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(2)
'Excelapp.ActiveSheet.Cells(2, 7) = DT1.Recordset.Fields(13)
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8)
DH = DT1.Recordset.Fields(1)




''''''''''''''''''''''''''''小卡
'Excelapp.ActiveSheet.Cells(22, 2) = dt1.Recordset.Fields(0)
'Excelapp.ActiveSheet.Cells(22, 5) = dt1.Recordset.Fields(14)
'Excelapp.ActiveSheet.Cells(22, 7) = dt1.Recordset.Fields(15)
'Excelapp.ActiveSheet.Cells(23, 1) = dt1.Recordset.Fields(8)
'Excelapp.ActiveSheet.Cells(23, 5) = dt1.Recordset.Fields(2)

'Excelapp.ActiveSheet.Cells(27, 2) = dt1.Recordset.Fields(3)
'Excelapp.ActiveSheet.Cells(31, 2) = dt1.Recordset.Fields(9) + Space(5) + "幅宽:" + dt1.Recordset.Fields(5) + "   克重" + dt1.Recordset.Fields(10)
Excelapp.ActiveSheet.Cells(5, 9) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '锅号条码

DT1.RecordSource = "select * from sczy_z where 单号='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)                ''''''''''总备注
End If


'dt1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
'dt1.Refresh

'If Not dt1.Recordset.EOF Then
'dt1.Recordset.MoveFirst
'i = 25
'Do While Not dt1.Recordset.EOF
'Excelapp.ActiveSheet.Cells(i, 1) = dt1.Recordset.Fields(3)
'Excelapp.ActiveSheet.Cells(i, 5) = dt1.Recordset.Fields(6)
'Excelapp.ActiveSheet.Cells(i, 6) = dt1.Recordset.Fields(7)
'i = i + 1
'dt1.Recordset.MoveNext
'Loop
'End If

DT1.RecordSource = "select round(SUM(重量),2),round(SUM(匹数),1) from kpd where 锅号='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
End If

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
DT1.Refresh
i = 0
L = ""
ZM = ""
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(10 + i * 2, 1) = DT1.Recordset.Fields(13)
Excelapp.ActiveSheet.Cells(10 + i * 2, 2) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(10 + i * 2, 4) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(10 + i * 2, 5) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(10 + i * 2, 6) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(7))
Excelapp.ActiveSheet.Cells(10 + i * 2, 7) = DT1.Recordset.Fields(19)
Excelapp.ActiveSheet.Cells(10 + i * 2, 8) = DT1.Recordset.Fields(10)
Excelapp.ActiveSheet.Cells(10 + i * 2, 9) = DT1.Recordset.Fields(30)     ''''''加工项目
Excelapp.ActiveSheet.Cells(10 + i * 2, 12) = DT1.Recordset.Fields(9)     '''备注
'If InStr(ZM, Trim(dt1.Recordset.Fields(30))) = 0 Then
'ZM = ZM + Trim(dt1.Recordset.Fields(30))
'End If
'L = L + Trim(dt1.Recordset.Fields(6)) + "+"
i = i + 1
DT1.Recordset.MoveNext
Loop
'Excelapp.ActiveSheet.Cells(5, 2) = ZM
'Excelapp.ActiveSheet.Cells(3, 9) = Mid(L, 1, Len(L) - 1)
'Excelapp.ActiveSheet.Cells(29, 2) = Mid(L, 1, Len(L) - 1)
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

Public Sub lcd22(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\流程单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "select 单号,锅号,品名,光胚幅宽,重量,色别,备注,日期,投染类别,面料用途,合同负责,下单日期,合同交期,总备注,计划日期,成分,技术要求,货号,色名,坯布类型,流程,标签,匹数 from v_kpd_mx where 锅号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(1, 9) = "*" + DT1.Recordset.Fields(1) + "J*" ''''锅号条码
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(8)   ''''加工类型
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(10)  '''业务
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(7))   '''日期
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(19)   '''坯布类型
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(15)   ''''成分
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(9)    '''面料用途
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(0)   '''合同号
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)   '''货号
Excelapp.ActiveSheet.Cells(5, 9) = DT1.Recordset.Fields(5)  '颜色
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(1)   '''缸号
Excelapp.ActiveSheet.Cells(6, 6) = Trim(DT1.Recordset.Fields(12))    '''交期
Excelapp.ActiveSheet.Cells(6, 9) = DT1.Recordset.Fields(18)   '色号
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)   '''布类
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(21)   ''原机型 先款号
If Val(DT1.Recordset.Fields(22)) = 0 Then
Excelapp.ActiveSheet.Cells(7, 9) = "" '计划匹数
Else
Excelapp.ActiveSheet.Cells(7, 9) = DT1.Recordset.Fields(22)  '计划匹数
End If
Excelapp.ActiveSheet.Cells(7, 12) = ""  '实际匹数
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(3)   '''幅宽
Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(16)    '''克重
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(4)  '计划重量
Excelapp.ActiveSheet.Cells(8, 12) = "" '实际重量
Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(20)   ''''流程
Excelapp.ActiveSheet.Cells(13, 2) = DT1.Recordset.Fields(13)   ''''总备注
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(6)   ''''备注
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
Public Sub lcd22yh(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\印花流程单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.RecordSource = "select 单号,锅号,品名,光胚幅宽,重量,色别,备注,日期,投染类别,面料用途,合同负责,下单日期,合同交期,总备注,计划日期,成分,技术要求,货号,色名,坯布类型,流程,标签,匹数 from v_kpd_mx where 锅号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(1, 9) = "*" + DT1.Recordset.Fields(1) + "J*" ''''锅号条码
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(8)   ''''加工类型
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(10)  '''业务
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(7))   '''日期
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(19)   '''坯布类型
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(15)   ''''成分
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(9)    '''面料用途
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(0)   '''合同号
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)   '''货号
Excelapp.ActiveSheet.Cells(5, 9) = DT1.Recordset.Fields(5)  '颜色
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(1)   '''缸号
Excelapp.ActiveSheet.Cells(6, 6) = Trim(DT1.Recordset.Fields(12))    '''交期
Excelapp.ActiveSheet.Cells(6, 9) = DT1.Recordset.Fields(18)   '色号
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)   '''布类
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(21)   ''原机型 先款号
If Val(DT1.Recordset.Fields(22)) = 0 Then
Excelapp.ActiveSheet.Cells(7, 9) = ""  '计划匹数
Else
Excelapp.ActiveSheet.Cells(7, 9) = DT1.Recordset.Fields(22)  '计划匹数
End If
Excelapp.ActiveSheet.Cells(7, 12) = ""  '实际匹数
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(3)   '''幅宽
Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(16)    '''克重
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(4)  '计划重量
Excelapp.ActiveSheet.Cells(8, 12) = "" '实际重量
Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(20)   ''''流程
Excelapp.ActiveSheet.Cells(13, 2) = DT1.Recordset.Fields(13)   ''''总备注
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(6)   ''''备注
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
Public Sub lcd22fx(DT1 As Adodc, gh As String, lb As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application
Dim b As Integer
Dim DH As String
On Error Resume Next

'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\流程单辅.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")
DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & a & "' and 卡号='" & lb & "'"
DT1.Refresh

b = DT1.Recordset.Fields(11)   ''''序号
DH = DT1.Recordset.Fields(1)   ''''单号

Excelapp.ActiveSheet.Cells(1, 2) = DT1.Recordset.Fields(51)    ''花型  '''是织布要求
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)    ''客户
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(14)   ''机台
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(3)    ''品名
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8) + DT1.Recordset.Fields(52) ''色别   颜色+色号
Excelapp.ActiveSheet.Cells(2, 10) = Trim(DT1.Recordset.Fields(12))   ''''日期
Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(2)        '''''锅号
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(13)        ''''款号
Excelapp.ActiveSheet.Cells(10, 3) = DT1.Recordset.Fields(5)        ''''光坯幅宽
Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(10)        ''''克重
Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(9)        ''''备注    染色要求  总备注
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(6))       ''''计划匹
Excelapp.ActiveSheet.Cells(4, 9) = Trim(DT1.Recordset.Fields(7))        ''''计划重
Excelapp.ActiveSheet.Cells(5, 9) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '锅号条码
Excelapp.ActiveSheet.Cells(12, 13) = DT1.Recordset.Fields(46)    ''  '''印花图案


''''''''''''''''''''''''''''

DT1.RecordSource = "select 备注,流程 as 总备注,花型 as 货号,成分 as 批号,缩水率 as 缩水,扭度 as 手感,布纹 as 说明 from sczy_x where 单号='" & DH & "' and 序号='" & b & "'"   '''货号 织布、 成分--批号 哪天来的
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(2)                ''''''''''织布
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(3)                ''''''''''批号
Excelapp.ActiveSheet.Cells(10, 5) = DT1.Recordset.Fields(0)                ''''''''''备注
'Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(1)                ''''''''''总备注
Excelapp.ActiveSheet.Cells(31, 2) = DT1.Recordset.Fields(4)                ''''''''''缩水率
Excelapp.ActiveSheet.Cells(32, 2) = DT1.Recordset.Fields(5)                ''''''''''缩水率
Excelapp.ActiveSheet.Cells(31, 8) = DT1.Recordset.Fields(6)                ''''''''''说明
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


Public Sub mprkbqdy(DT1 As Adodc, gh As String, xh As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\mprkbq.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select 客户名称,布类,和约号,毛胚幅宽,毛胚匹数,毛胚重量,备注,负责人,日期,ny,存放位置 from ckgl where 单据号='" & gh & "' and ip='" & xh & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(1, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(1, 5) = Trim(DT1.Recordset.Fields(8))
Excelapp.ActiveSheet.Cells(2, 2) = Trim(DT1.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(9)
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(10)
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

Public Sub wtlcd22(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\委外流程单多ok.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = cast('" & a & "' as real)"
DT1.Refresh
'Excelapp.ActiveSheet.Cells(2, 2) = dt1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(2, 4) = DT1.Recordset.Fields(15)
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(14)
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(12))
Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(2)
'Excelapp.ActiveSheet.Cells(2, 7) = DT1.Recordset.Fields(13)
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8)
DH = DT1.Recordset.Fields(1)




''''''''''''''''''''''''''''
'Excelapp.ActiveSheet.Cells(22, 2) = dt1.Recordset.Fields(0)
'Excelapp.ActiveSheet.Cells(22, 5) = dt1.Recordset.Fields(14)
'Excelapp.ActiveSheet.Cells(22, 7) = dt1.Recordset.Fields(15)
'Excelapp.ActiveSheet.Cells(23, 1) = dt1.Recordset.Fields(8)
'Excelapp.ActiveSheet.Cells(23, 5) = dt1.Recordset.Fields(2)

'Excelapp.ActiveSheet.Cells(27, 2) = dt1.Recordset.Fields(3)
'Excelapp.ActiveSheet.Cells(31, 2) = dt1.Recordset.Fields(9) + Space(5) + "幅宽:" + dt1.Recordset.Fields(5) + "   克重" + dt1.Recordset.Fields(10)
'Excelapp.ActiveSheet.Cells(5, 9) = "*" + dt1.Recordset.Fields(2) + "J" + "*"  '锅号条码

DT1.RecordSource = "select * from sczy_z where 单号='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)                ''''''''''总备注
End If

DT1.RecordSource = "select * from kpdwwjg where 委外锅号='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)                ''''''''''委外单位
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(2)                ''''''''''委外信息
End If


'dt1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
'dt1.Refresh

'If Not dt1.Recordset.EOF Then
'dt1.Recordset.MoveFirst
'i = 25
'Do While Not dt1.Recordset.EOF
'Excelapp.ActiveSheet.Cells(i, 1) = dt1.Recordset.Fields(3)
'Excelapp.ActiveSheet.Cells(i, 5) = dt1.Recordset.Fields(6)
'Excelapp.ActiveSheet.Cells(i, 6) = dt1.Recordset.Fields(7)
'i = i + 1
'dt1.Recordset.MoveNext
'Loop
'End If

DT1.RecordSource = "select round(SUM(重量),2),round(SUM(匹数),1) from kpd where 锅号='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
End If

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' order by IP"
DT1.Refresh
i = 0
L = ""
ZM = ""
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(10 + i * 2, 1) = DT1.Recordset.Fields(13)
Excelapp.ActiveSheet.Cells(10 + i * 2, 2) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(10 + i * 2, 4) = DT1.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(10 + i * 2, 5) = DT1.Recordset.Fields(5)
Excelapp.ActiveSheet.Cells(10 + i * 2, 6) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(7))
Excelapp.ActiveSheet.Cells(10 + i * 2, 7) = DT1.Recordset.Fields(19)
Excelapp.ActiveSheet.Cells(10 + i * 2, 8) = DT1.Recordset.Fields(10)
Excelapp.ActiveSheet.Cells(10 + i * 2, 9) = Trim(DT1.Recordset.Fields(30)) ''''''加工项目
Excelapp.ActiveSheet.Cells(10 + i * 2, 12) = DT1.Recordset.Fields(9)  '''''+备注
'If InStr(ZM, Trim(dt1.Recordset.Fields(30))) = 0 Then
'ZM = ZM + Trim(dt1.Recordset.Fields(30))
'End If
'L = L + Trim(dt1.Recordset.Fields(6)) + "+"
i = i + 1
DT1.Recordset.MoveNext
Loop
'Excelapp.ActiveSheet.Cells(5, 2) = ZM
'Excelapp.ActiveSheet.Cells(3, 9) = Mid(L, 1, Len(L) - 1)
'Excelapp.ActiveSheet.Cells(29, 2) = Mid(L, 1, Len(L) - 1)
Excelapp.Visible = True
    Excelapp.DisplayAlerts = False
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.Quit
    Set Excelapp = Nothing
    Exit Sub

errorhandler:
    MsgBox "Error: " & Err.Description
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
End Sub
Public Sub lcd22f(DT1 As Adodc, dt2 As Adodc, gh As String, lb As String, selectedPrinter As String)
    Dim Excelapp As Object ' 使用通用对象声明
    Dim Workbook As Object
    Dim Worksheet As Object
    
    On Error GoTo Ert

    ' 创建 Excel 应用程序对象
    Set Excelapp = CreateObject("Excel.Application")
    If Excelapp Is Nothing Then
        MsgBox "无法创建 Excel 应用程序对象。请确保已安装 Excel，并检查权限和注册表设置。"
        Exit Sub
    End If

    ' 设置 Excel 应用程序属性
    Excelapp.Caption = "永发打印模版软件之打印"
    Excelapp.SheetsInNewWorkbook = 1

    ' 打开已有的工作簿
    Set Workbook = Excelapp.Workbooks.Open(App.Path & "\打印模版\广兴\生产流程卡.xls")
    Set Worksheet = Workbook.Sheets(1)
    Worksheet.Activate

    ' 设置数据库连接和查询
    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
    DT1.Refresh
    If DT1.Recordset.EOF Then GoTo Cleanup

    Dim maxWeight As Double
    maxWeight = DT1.Recordset.Fields("zl")

    DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & maxWeight & "' and 卡号='" & lb & "'"
    DT1.Refresh
    If DT1.Recordset.EOF Then GoTo Cleanup


'Excelapp.ActiveSheet.Cells(3, 3) = Trim(lb)   ''''卡号
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(12) '''备布卡日期
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(2) '''锅号
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''客户
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(3) '''布类
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8) '''颜色
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(52) '''色号
Excelapp.ActiveSheet.Cells(9, 2) = DT1.Recordset.Fields(9) '''染色要求

Excelapp.ActiveSheet.Cells(2, 8) = DT1.Recordset.Fields(0) '''排缸卡客户
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 10) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(52) ''色号
Excelapp.ActiveSheet.Cells(5, 8) = DT1.Recordset.Fields(5) ''幅宽
Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(10) ''克重
Excelapp.ActiveSheet.Cells(7, 7) = DT1.Recordset.Fields(9) '''染色要求
Excelapp.ActiveSheet.Cells(11, 7) = DT1.Recordset.Fields(12) '''排缸卡日期


Excelapp.ActiveSheet.Cells(21, 2) = DT1.Recordset.Fields(0)   ''''客户
Excelapp.ActiveSheet.Cells(18, 6) = "*" + DT1.Recordset.Fields(2) + "J*" '''条码
Excelapp.ActiveSheet.Cells(20, 5) = DT1.Recordset.Fields(2)   '''锅号

Excelapp.ActiveSheet.Cells(20, 2) = Trim(DT1.Recordset.Fields(12))    ''''日期
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(18, 1) = "返修"
Else
Excelapp.ActiveSheet.Cells(18, 1) = "正常"
End If '''' 类别
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''款号
Excelapp.ActiveSheet.Cells(23, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''色号+颜色
Excelapp.ActiveSheet.Cells(24, 2) = DT1.Recordset.Fields(9)     '''''染色要求
Excelapp.ActiveSheet.Cells(22, 2) = DT1.Recordset.Fields(3)   ''品名
Excelapp.ActiveSheet.Cells(26, 2) = DT1.Recordset.Fields(5)     ''幅宽
 Excelapp.ActiveSheet.Cells(23, 6) = DT1.Recordset.Fields(82) ''幅宽明细
Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(10)  ''克重
''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(重量,0)),2),SUM(isnull(匹数,0)) from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(1)   ''''备布卡匹数
Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(0)    ''''备布卡计划量
Excelapp.ActiveSheet.Cells(4, 10) = DT1.Recordset.Fields(1)   ''''排缸卡匹数
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(0)    ''''排缸卡计划量

Excelapp.ActiveSheet.Cells(28, 2) = DT1.Recordset.Fields(1)   ''''匹数
Excelapp.ActiveSheet.Cells(29, 2) = DT1.Recordset.Fields(0)    ''''计划量
End If

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "'  order by 卡号,IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(16 + i * 1, 2) = DT1.Recordset.Fields(55)   '''编号
Excelapp.ActiveSheet.Cells(16 + i * 1, 3) = DT1.Recordset.Fields(3)   '''品名
Excelapp.ActiveSheet.Cells(16 + i * 1, 5) = DT1.Recordset.Fields(5)   '''幅宽
Excelapp.ActiveSheet.Cells(16 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''克重
Excelapp.ActiveSheet.Cells(16 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''匹数
Excelapp.ActiveSheet.Cells(16 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''重量
Excelapp.ActiveSheet.Cells(16 + i * 1, 9) = DT1.Recordset.Fields("卡号")       ''''卡号

i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct 编号,mr from kpd where  锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''流程
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(31, 2) = lc  ''''流程

'''用数组将流程分开并竖着打印在表格上
'Dim dataArray() As String
'dataArray = Split(lc, "-")

'Dim L As Integer
'For L = 0 To UBound(dataArray)
'   Excelapp.ActiveSheet.Cells(L + 38, 1).value = dataArray(L)
'Next L


'DT1.RecordSource = "select distinct 编号,备注 from kpd where  锅号='" & gh & "' and 卡号='" & lb & "'"
'DT1.Refresh
'If Not DT1.Recordset.EOF Then
'bz = ""
'xbz = ""
'Do While Not DT1.Recordset.EOF

'If InStr(xbz, DT1.Recordset.Fields(1)) = 0 Then
'xbz = xbz + DT1.Recordset.Fields(1)
'End If

dt2.RecordSource = "select * from ckgl where 单据号='" & gh & "'"     ''这里单据号必须等于gh才能调出业务
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If
Excelapp.ActiveSheet.Cells(5, 2) = dt2.Recordset.Fields(12) ''备布卡来料单位
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(9) '''备布卡存放位置

Excelapp.ActiveSheet.Cells(20, 8) = dt2.Recordset.Fields(16)  ''司机业务
Excelapp.ActiveSheet.Cells(21, 8) = dt2.Recordset.Fields(12) ''来料单位
Excelapp.ActiveSheet.Cells(30, 2) = dt2.Recordset.Fields(9) '''存放位置
End If


    ' 设置打印属性并打印
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.DisplayAlerts = False

    ' 切换到用户选择的打印机
    If selectedPrinter <> "" Then
        TrySetActivePrinter Excelapp, selectedPrinter
    End If

    ' 打印工作表
    Worksheet.PrintOut Copies:=1, Preview:=False, PrintToFile:=False, Collate:=True

Cleanup:
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
    Exit Sub

Ert:
    MsgBox "An error occurred: " & Err.Description
    Resume Cleanup
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
Public Sub lcd222f(DT1 As Adodc, dt2 As Adodc, gh As String, lb As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")
Dim lc As String


Excelapp.Caption = "广兴染整软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\生产流程卡.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & a & "' and 卡号='" & lb & "'"
DT1.Refresh
'Excelapp.ActiveSheet.Cells(3, 3) = Trim(lb)   ''''卡号
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(12) '''备布卡日期
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(2) '''锅号
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''客户
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(3) '''布类
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8) '''颜色
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(52) '''色号
Excelapp.ActiveSheet.Cells(9, 2) = DT1.Recordset.Fields(9) '''染色要求

Excelapp.ActiveSheet.Cells(2, 8) = DT1.Recordset.Fields(0) '''排缸卡客户
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 10) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(52) ''色号
Excelapp.ActiveSheet.Cells(5, 8) = DT1.Recordset.Fields(5) ''幅宽
Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(10) ''克重
Excelapp.ActiveSheet.Cells(7, 7) = DT1.Recordset.Fields(9) '''染色要求
Excelapp.ActiveSheet.Cells(11, 7) = DT1.Recordset.Fields(12) '''排缸卡日期


Excelapp.ActiveSheet.Cells(21, 2) = DT1.Recordset.Fields(0)   ''''客户
Excelapp.ActiveSheet.Cells(18, 6) = "*" + DT1.Recordset.Fields(2) + "J*" '''条码
Excelapp.ActiveSheet.Cells(20, 5) = DT1.Recordset.Fields(2)   '''锅号

Excelapp.ActiveSheet.Cells(20, 2) = Trim(DT1.Recordset.Fields(12))    ''''日期
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(18, 1) = "返修"
Else
Excelapp.ActiveSheet.Cells(18, 1) = "正常"
End If '''' 类别
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''款号
Excelapp.ActiveSheet.Cells(23, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''色号+颜色
Excelapp.ActiveSheet.Cells(24, 2) = DT1.Recordset.Fields(9)     '''''染色要求
Excelapp.ActiveSheet.Cells(22, 2) = DT1.Recordset.Fields(3)   ''品名
Excelapp.ActiveSheet.Cells(26, 2) = DT1.Recordset.Fields(5)     ''幅宽
 Excelapp.ActiveSheet.Cells(23, 6) = DT1.Recordset.Fields(82) ''幅宽明细
Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(10)  ''克重
''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(重量,0)),2),SUM(isnull(匹数,0)) from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(1)   ''''备布卡匹数
Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(0)    ''''备布卡计划量
Excelapp.ActiveSheet.Cells(4, 10) = DT1.Recordset.Fields(1)   ''''排缸卡匹数
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(0)    ''''排缸卡计划量

Excelapp.ActiveSheet.Cells(28, 2) = DT1.Recordset.Fields(1)   ''''匹数
Excelapp.ActiveSheet.Cells(29, 2) = DT1.Recordset.Fields(0)    ''''计划量
End If

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "'  order by 卡号,IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(16 + i * 1, 2) = DT1.Recordset.Fields(55)   '''编号
Excelapp.ActiveSheet.Cells(16 + i * 1, 3) = DT1.Recordset.Fields(3)   '''品名
Excelapp.ActiveSheet.Cells(16 + i * 1, 5) = DT1.Recordset.Fields(5)   '''幅宽
Excelapp.ActiveSheet.Cells(16 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''克重
Excelapp.ActiveSheet.Cells(16 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''匹数
Excelapp.ActiveSheet.Cells(16 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''重量
Excelapp.ActiveSheet.Cells(16 + i * 1, 9) = DT1.Recordset.Fields("卡号")       ''''卡号

i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct 编号,mr from kpd where  锅号='" & gh & "' and 卡号='" & lb & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''流程
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(31, 2) = lc  ''''流程

'''用数组将流程分开并竖着打印在表格上
'Dim dataArray() As String
'dataArray = Split(lc, "-")

'Dim L As Integer
'For L = 0 To UBound(dataArray)
'   Excelapp.ActiveSheet.Cells(L + 38, 1).value = dataArray(L)
'Next L


'DT1.RecordSource = "select distinct 编号,备注 from kpd where  锅号='" & gh & "' and 卡号='" & lb & "'"
'DT1.Refresh
'If Not DT1.Recordset.EOF Then
'bz = ""
'xbz = ""
'Do While Not DT1.Recordset.EOF

'If InStr(xbz, DT1.Recordset.Fields(1)) = 0 Then
'xbz = xbz + DT1.Recordset.Fields(1)
'End If

dt2.RecordSource = "select * from ckgl where 单据号='" & gh & "'"     ''这里单据号必须等于gh才能调出业务
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If
Excelapp.ActiveSheet.Cells(5, 2) = dt2.Recordset.Fields(12) ''备布卡来料单位
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(9) '''备布卡存放位置

Excelapp.ActiveSheet.Cells(20, 8) = dt2.Recordset.Fields(16)  ''司机业务
Excelapp.ActiveSheet.Cells(21, 8) = dt2.Recordset.Fields(12) ''来料单位
Excelapp.ActiveSheet.Cells(30, 2) = dt2.Recordset.Fields(9) '''存放位置
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
Public Sub lcd22f2(DT1 As Adodc, dt2 As Adodc, gh As String, lb As String)
    Dim Excelapp As Excel.Application
    Set Excelapp = New Excel.Application
    Excelapp.Visible = False  ' Initially hide the application to prevent screen flickering

    On Error GoTo errorhandler

    ' Open the Excel template
    Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\流程卡ok.xls")
    Excelapp.Sheets(1).Activate
    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "' and 卡号='" & lb & "'"
    DT1.Refresh
    Dim maxWeight As Variant
    maxWeight = DT1.Recordset.Fields("zl").value

    DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & maxWeight & "' and 卡号='" & lb & "'"
    DT1.Refresh

    With Excelapp.ActiveSheet
        .Cells(3, 9).value = DT1.Recordset.Fields(2).value ' 锅号
        .Cells(3, 2).value = DT1.Recordset.Fields(0).value ' 客户
        .Cells(4, 2).value = DT1.Recordset.Fields(3).value ' 布类
        .Cells(3, 6).value = DT1.Recordset.Fields(8).value ' 颜色
        .Cells(4, 6).value = DT1.Recordset.Fields(52).value ' 色号
        .Cells(6, 3).value = DT1.Recordset.Fields(9).value ' 染色要求
        .Cells(5, 2).value = DT1.Recordset.Fields(13).value ' 款号
        .Cells(8, 2).value = DT1.Recordset.Fields(5).value ' 幅宽
        .Cells(8, 6).value = DT1.Recordset.Fields(10).value ' 克重
        .Cells(1, 7) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '锅号条码
  ' 获取幅宽明细数据
    Dim widthDetails As String
    widthDetails = DT1.Recordset.Fields("幅宽明细").value
    Dim items() As String
    items = Split(widthDetails, " ")
    
    ' 定义起始行及列
    Dim startRow As Integer, column As Integer
    startRow = 11    ' 从第11行开始
    column = 2      ' 从第2列开始

    ' 定义当前列已打印的行数
    Dim printedRows As Integer
    printedRows = 0

    ' 遍历数据项
    For i = LBound(items) To UBound(items)
        ' 判断是否包含数字+cm，若是则加粗显示，并设置字体大小为18号
        If InStr(items(i), "cm") > 0 Then
            .Cells(startRow + printedRows, column).value = items(i)
            .Cells(startRow + printedRows, column).Font.Bold = True ' 加粗显示
            .Cells(startRow + printedRows, column).Font.Size = 18 ' 设置字体大小为18号
        ElseIf InStr(items(i), "领") > 0 Or InStr(items(i), "袖") > 0 Then
            .Cells(startRow + printedRows, column).value = items(i)
            .Cells(startRow + printedRows, column).Font.Bold = True ' 加粗显示
            .Cells(startRow + printedRows, column).Font.Size = 18 ' 设置字体大小为18号
        Else
            .Cells(startRow + printedRows, column).value = items(i)
            ' 针对不是数字+cm的情况，取消加粗显示，并设置字体大小为18号
            .Cells(startRow + printedRows, column).Font.Bold = False
            .Cells(startRow + printedRows, column).Font.Size = 18 ' 设置字体大小为18号
        End If
        
        .Cells(startRow + printedRows, column).WrapText = False  ' 禁用换行

        ' 当打印的行数超过26时，切换到下一列并重置打印行数
        printedRows = printedRows + 1
        If printedRows >= 26 Then
            column = column + 1    ' 切换到下一列
            printedRows = 0  ' 重置打印行数
        End If
    Next i
End With

DT1.RecordSource = "select round(SUM(isnull(重量,0)),2),SUM(isnull(匹数,0)) from kpd where 锅号='" & gh & "' "
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(1) & "匹" ' 总匹数
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(0) & "kg" ' 总重量
End If
dt2.RecordSource = "select * from ckgl where 单据号='" & gh & "'"     ''这里单据号必须等于gh才能调出业务
dt2.Refresh
If Not dt2.Recordset.EOF Then

Excelapp.ActiveSheet.Cells(5, 9) = dt2.Recordset.Fields(16)  ''司机业务
Excelapp.ActiveSheet.Cells(4, 9) = dt2.Recordset.Fields(12) ''来料单位

End If

 Excelapp.Visible = True
    Excelapp.DisplayAlerts = False
    Excelapp.ActiveWindow.Zoom = 100
    Exit Sub
    
errorhandler:
    MsgBox "Error: " & Err.Description
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
End Sub
Public Sub mpckdy(DT1 As Adodc, gh As String, kh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next


Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\毛坯配缸.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select 缸号,布类,sum(毛胚匹数),round(sum(毛胚重量),2),出库日期 from mpbh where 锅号='" & gh & "' and 缸号 in(select distinct 编号 from kpd where 锅号='" & gh & "' and 卡号='" & kh & "') group by 缸号,布类,出库日期 order by 缸号"
DT1.Refresh
L = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(31, 6) = Trim(DT1.Recordset.Fields(4))   ''''出库日期
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(33 + L, 2) = DT1.Recordset.Fields(0)   ''''缸号
Excelapp.ActiveSheet.Cells(33 + L, 4) = DT1.Recordset.Fields(1)   ''布类
Excelapp.ActiveSheet.Cells(33 + L, 6) = Trim(DT1.Recordset.Fields(3))   '''''重量
Excelapp.ActiveSheet.Cells(33 + L, 7) = Trim(DT1.Recordset.Fields(2))   '''''匹数
L = L + 1
DT1.Recordset.MoveNext
Loop
End If
If kh = "甲" Then
Excelapp.ActiveSheet.Cells(38, 2) = "主料"  '''''合计匹数
Else
Excelapp.ActiveSheet.Cells(38, 2) = "辅料"  '''''合计匹数
End If

DT1.RecordSource = "select isnull(sum(毛胚匹数),0),isnull(round(sum(毛胚重量),2),0) from mpbh where 锅号='" & gh & "'"
DT1.Refresh
Excelapp.ActiveSheet.Cells(38, 2) = "合计"  '''''合计匹数
Excelapp.ActiveSheet.Cells(38, 7) = Trim(DT1.Recordset.Fields(0))   '''''合计匹数
Excelapp.ActiveSheet.Cells(38, 6) = Trim(DT1.Recordset.Fields(1))   '''''合计重量

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
Function ExtractWidthAndWeights(Data As String) As Collection
    Dim regex As New RegExp
    Dim matches As MatchCollection
    Dim widthDetails As New Collection

    ' 正则表达式匹配幅宽数字后跟"cm"和后面的数字，或单独的"领"或"袖"和后面的数字
    regex.Pattern = "(\d+\s*cm\s*(\d+\s*)+)|(领\s*(\d+\s*)+)|(袖\s*(\d+\s*)+)"
    regex.Global = True
    regex.IgnoreCase = True

    ' 执行匹配
    Set matches = regex.Execute(Data)

    Dim match As match
    For Each match In matches
        ' 创建一个新的集合来保存幅宽和对应的重量
        Dim itemCollection As New Collection
        Dim content As String
        content = match.value

        ' 使用空格分割幅宽和重量
        Dim parts() As String
        parts = Split(content, " ")

        Dim i As Integer
        For i = 1 To UBound(parts)
            If IsNumeric(parts(i)) Then
                itemCollection.Add parts(i) ' 添加重量到集合
            End If
        Next

        widthDetails.Add itemCollection, parts(0) ' 使用幅宽作为键
    Next

    Set ExtractWidthAndWeights = widthDetails
End Function

Public Sub lcd22f3(DT1 As Adodc, dt2 As Adodc, gh As String, weight As Double, count As Double)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")
Dim lc As String


Excelapp.Caption = "广兴打印模版软件之打印"
'3)添加新工作簿：
'4)打开已存在的工作簿：
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\出库锅单.xls")
'5)设置第2个工作表为活动工作表：
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(重量) as zl from kpd where 锅号='" & gh & "' "
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where 锅号='" & gh & "' And 重量 = '" & a & "' "
DT1.Refresh


Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   ''''客户
Excelapp.ActiveSheet.Cells(1, 5) = DT1.Recordset.Fields(2)   '''锅号

Excelapp.ActiveSheet.Cells(1, 2) = Trim(DT1.Recordset.Fields(12))    ''''日期

Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''色号+颜色
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(9)     '''''染色要求
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(3)   ''品名
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(5)     ''幅宽
 Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(82) ''幅宽明细
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(10)  ''克重
Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(83) ''布头


''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(重量,0)),2),SUM(isnull(匹数,0)) from kpd where 锅号='" & gh & "' "
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else

Excelapp.ActiveSheet.Cells(9, 2).value = count   ' 匹数
Excelapp.ActiveSheet.Cells(10, 2).value = weight ' 计划量

End If
DT1.RecordSource = "select distinct 编号,mr from kpd where  锅号='" & gh & "' "
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''流程
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(12, 2) = lc  ''''流程


dt2.RecordSource = "select * from ckgl where 单据号='" & gh & "'"     ''这里单据号必须等于gh才能调出业务
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If

Excelapp.ActiveSheet.Cells(8, 5) = dt2.Recordset.Fields(18) ''大布重量
Excelapp.ActiveSheet.Cells(9, 5) = dt2.Recordset.Fields(21) ''大布匹数
Excelapp.ActiveSheet.Cells(10, 5) = dt2.Recordset.Fields(20) ''领子匹数
Excelapp.ActiveSheet.Cells(11, 5) = dt2.Recordset.Fields(19) ''领子重量

Excelapp.ActiveSheet.Cells(1, 8) = dt2.Recordset.Fields(16)  ''司机业务
Excelapp.ActiveSheet.Cells(2, 8) = dt2.Recordset.Fields(12) ''来料单位
Excelapp.ActiveSheet.Cells(11, 2) = dt2.Recordset.Fields(9) '''存放位置
End If

Excelapp.ActiveWindow.Zoom = 100   ' 设置窗口缩放比例为 100%
     'Excelapp.Visible = True  ' 注释掉设置 Excel 应用程序可见的代码
    Excelapp.DisplayAlerts = False   ' 禁用显示警告

    Excelapp.ActiveSheet.PrintOut   ' 直接打印当前工作表

    Set Excelapp = Nothing   ' 释放 Excel 应用程序对象
    Exit Sub   ' 退出子程序

Ert:

'Excelapp.Quit '关闭EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
