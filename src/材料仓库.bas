Attribute VB_Name = "���ϲֿ�"
Public Sub clrk(DT1 As Adodc, DH As String)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\clrk.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT ��Ӧ��λ,��������,���Ϲ��,���ϵ�λ,��ɫ,����,����,�ϼƽ��,����,��ע,����,����  FROM clgl WHERE ���ݺ�='" & DH & "' order BY ���"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select ���� from gys where ���='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 15) = Trim(DH)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(8))
'        Excelapp.ActiveSheet.Cells(13, 4) = "ҵ��Ա��" + DT3.Recordset.Fields(0)
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

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub



Public Sub clck(DT1 As Adodc, DH As String)  ''''�ޱ���

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


       Excelapp.SheetsInNewWorkbook = 10

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\clck.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT ���ϳ���,��������,���Ϲ��,��ɫ,����,���ϵ�λ,����,����,�ϼƽ��,����,��ע  FROM clkpd WHERE ���ݺ�='" & DH & "' order BY ���"
DT1.Refresh

DT1.Recordset.MoveFirst
'DT3.RecordSource = "select ���� from gys where ���='" & DT1.Recordset.Fields(0) & "'"
'DT3.Refresh
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
        Excelapp.ActiveSheet.Cells(3, 6) = Trim(DT1.Recordset.Fields(9))
        Excelapp.ActiveSheet.Cells(3, 14) = Trim(DH)
'        Excelapp.ActiveSheet.Cells(13, 4) = "ҵ��Ա��" + DT3.Recordset.Fields(0)
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

'Excelapp.Quit '�ر�EXCEL
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\cbfx.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

