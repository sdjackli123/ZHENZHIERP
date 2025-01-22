Attribute VB_Name = "材料仓库"
Public Sub clrk(DT1 As Adodc, DH As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\clrk.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT 供应单位,材料名称,材料规格,材料单位,颜色,数量,单价,合计金额,日期,备注,批次,包件  FROM clgl WHERE 单据号='" & DH & "' order BY 序号"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 15) = Trim(DH)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(8))
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(10)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(11)
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 8) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 10) = DT1.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(i, 12) = Format(DT1.Recordset.Fields(7), "#0.00")
        Excelapp.ActiveSheet.Cells(i, 15) = DT1.Recordset.Fields(9)
       
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



Public Sub clck(DT1 As Adodc, DH As String)  ''''无标题

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\clck.xls")
'5)设置第2个工作表为活动工作表：
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT 领料车间,材料名称,材料规格,颜色,批次,材料单位,数量,单价,合计金额,日期,备注  FROM clkpd WHERE 单据号='" & DH & "' order BY 序号"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select 代码 from gys where 简称='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(9))
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "业务员：" + DT3.Recordset.Fields(0)
i = 6
Do While Not DT1.Recordset.EOF
        Excelapp.ActiveSheet.Cells(i, 1) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(i, 3) = DT1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(i, 4) = DT1.Recordset.Fields(3)
        Excelapp.ActiveSheet.Cells(i, 5) = DT1.Recordset.Fields(4)
        Excelapp.ActiveSheet.Cells(i, 6) = DT1.Recordset.Fields(5)
        Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(6)
        Excelapp.ActiveSheet.Cells(i, 9) = DT1.Recordset.Fields(7)
        Excelapp.ActiveSheet.Cells(i, 11) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(10)

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

Public Sub MXCBFX(Flex As VSFlexGrid, BT As String)

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
Excelapp.Workbooks.Open (App.Path & "\打印模版\广兴\cbfx.xls")
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

