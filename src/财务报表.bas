Attribute VB_Name = "���񱨱�"
Public Sub YEBDOutadodcToExcelSZ(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ���

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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\YEB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(3).Activate
DT1.RecordSource = "SELECT ��ƿ�Ŀ FROM ZLCX  GROUP BY ��ƿ�Ŀ"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT ��ƿ�Ŀ,������� FROM ZLCX  GROUP BY ��ƿ�Ŀ,�������"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''�ƶ�����һ����¼
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "ҳ"   '''''ҳ
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''�ڼ�



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "ҳ" '''''ҳ
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''�ڼ�
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  ��ƿ�Ŀ='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "����" Then
Excelapp.ActiveSheet.Cells(k, 4) = Val(dt2.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(k, 5) = Val(dt2.Recordset.Fields(5))
End If
       
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "3" Then
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

Public Sub ZYEBDOutadodcToExcelSZ(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ���

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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\YEB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(1).Activate
DT1.RecordSource = "SELECT ��ƿ�Ŀ FROM ZLCX  GROUP BY ��ƿ�Ŀ"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT ��ƿ�Ŀ,������� FROM ZLCX  GROUP BY ��ƿ�Ŀ,�������"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''�ƶ�����һ����¼
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "ҳ"   '''''ҳ
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''�ڼ�



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "ҳ" '''''ҳ
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''�ڼ�
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  ��ƿ�Ŀ='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "����" Then
Excelapp.ActiveSheet.Cells(k, 4) = Val(dt2.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(k, 5) = Val(dt2.Recordset.Fields(5))
End If
       
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "3" Then
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
Public Sub SYEBDOutadodcToExcelSZ(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ���

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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\YEB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(2).Activate
DT1.RecordSource = "SELECT ��ƿ�Ŀ FROM ZLCX  GROUP BY ��ƿ�Ŀ"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT ��ƿ�Ŀ,������� FROM ZLCX  GROUP BY ��ƿ�Ŀ,�������"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''�ƶ�����һ����¼
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "ҳ"   '''''ҳ
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''�ڼ�



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "ҳ" '''''ҳ
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''�ڼ�
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  ��ƿ�Ŀ='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "����" Then
Excelapp.ActiveSheet.Cells(k, 4) = Val(dt2.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(k, 5) = Val(dt2.Recordset.Fields(5))
End If
       
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "3" Then
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
Public Sub QYEBDOutadodcToExcelSZ(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ���

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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\YEB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(2).Activate
DT1.RecordSource = "SELECT ��ƿ�Ŀ FROM ZLCX  GROUP BY ��ƿ�Ŀ"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT ��ƿ�Ŀ,������� FROM ZLCX  GROUP BY ��ƿ�Ŀ,�������"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''�ƶ�����һ����¼
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "ҳ"   '''''ҳ
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''�ڼ�



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "ҳ" '''''ҳ
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''�ڼ�
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  ��ƿ�Ŀ='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "����" Then
Excelapp.ActiveSheet.Cells(k, 4) = Val(dt2.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(k, 5) = Val(dt2.Recordset.Fields(5))
End If
       
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "3" Then
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

Public Sub ZHZGZOutadodcToExcel(Flex As VSFlexGrid)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\GZB.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 2 To k

                          For j = 1 To 3

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 3, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 1 Then
 '        Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
'Excelapp.ActiveSheet.Cells(2, 3) = BT
Excelapp.ActiveWindow.SplitRow = 1  '��ֵ�һ��
Excelapp.ActiveWindow.SplitColumn = 0   '�����
Excelapp.ActiveWindow.FreezePanes = True   '�̶����


Excelapp.Cells.EntireColumn.AutoFit
'Excelapp.ActiveSheet.Columns(1).ColumnsWidth = 5
'Excelapp.ActiveSheet.Columns(2).ColumnsWidth = 125

 Excelapp.ActiveSheet.PageSetup.PrintTitleRows = "$1:$2"
       ' B.ҳ��:
Excelapp.ActiveSheet.PageSetup.CenterFooter = "��&Pҳ"
'C.ҳü�����˱߾�2cm:
Excelapp.ActiveSheet.PageSetup.HeaderMargin = 1 / 0.035
'D.ҳ�ŵ��׶˱߾�3cm:
Excelapp.ActiveSheet.PageSetup.HeaderMargin = 2 / 0.035
'e.���߾�2cm:
Excelapp.ActiveSheet.PageSetup.TopMargin = 1 / 0.035
'f.�ױ߾�2cm:
Excelapp.ActiveSheet.PageSetup.BottomMargin = 1 / 0.035
'g.��߾�2cm:
Excelapp.ActiveSheet.PageSetup.LeftMargin = 1 / 0.035
'h.�ұ߾�2cm:
Excelapp.ActiveSheet.PageSetup.RightMargin = 1 / 0.035
'i.ҳ��ˮƽ����:
Excelapp.ActiveSheet.PageSetup.CenterHorizontally = 1 / 0.035
'j.ҳ�洹ֱ����:
'Excelapp.ActiveSheet.PageSetup.CenterVertically = 2 / 0.035
'k.��ӡ��Ԫ������:
Excelapp.ActiveSheet.PageSetup.PrintGridlines = True

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

Public Sub XHGZOutadodcToExcel(Flex As VSFlexGrid)

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\GZB.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(2).Activate
        Q = 0

       ' Excelapp.Selection.Font.FontStyle = "Bold"

       ' Excelapp.Selection.FontSize = 6
        With Flex

                k = .Rows


     '''   Excelapp.ActiveSheet.Range("a3:" & Hang & .Rows + 2).Font.Size = 9            'xlBorderLineStyleContinuous

          For i = 2 To k

                          For j = 1 To 3

                              
                              DoEvents

                              Excelapp.ActiveSheet.Cells(i + 3, j) = "'" & .TextMatrix(i - 1, j)
                      
                          Next j
               
         If i >= 1 Then
 '        Q = Val(Excelapp.ActiveSheet.Cells(i, FD)) + Q
         End If
         Next i

        End With
'Excelapp.ActiveSheet.Cells(2, 3) = BT
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

Public Sub OutadodcToExcel11(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, fd9, fd10, fd11, BT) ''''��һ�ֶκϼƣ������⣩

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
        Q7 = 0
        Q8 = 0
        Q9 = 0
        Q10 = 0
        Q11 = 0

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
         End If
         Next i

        End With

Excelapp.ActiveSheet.Cells(1, 1) = BT

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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


Public Sub OutadodcToExcel8(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, BT) ''''��һ�ֶκϼƣ������⣩

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate
        Q1 = 0
        Q2 = 0
        Q3 = 0
        Q4 = 0
        Q5 = 0
        Q6 = 0
        Q7 = 0
        Q8 = 0

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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub


Public Sub FYEBDOutadodcToExcelSZ(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ��� �տ�����

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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\YEB.xls")
'5)���õ�1��������Ϊ�������
Excelapp.Sheets(2).Activate
DT1.RecordSource = "SELECT ��ƿ�Ŀ FROM ZLCX  GROUP BY ��ƿ�Ŀ"
DT1.Refresh
L = DT1.Recordset.RecordCount
If L < 1 Then Exit Sub

TT = L / 48
If TT <> Int(TT) Then
TT = Int(TT) + 1
End If

DT1.RecordSource = "SELECT ��ƿ�Ŀ,������� FROM ZLCX  GROUP BY ��ƿ�Ŀ,�������"
DT1.Refresh

DT1.Recordset.MoveFirst    ''''''''''�ƶ�����һ����¼
i = 1
k = 53 * (i - 1) + 5
Q = 1
Excelapp.ActiveSheet.Cells(2, 7) = Str(TT) + "-" + Str(i) + "ҳ"   '''''ҳ
Excelapp.ActiveSheet.Cells(2, 5) = QJ   '''''�ڼ�



Do While Not DT1.Recordset.EOF
If Q / 47 = Int(Q / 47) Then
i = i + 1
k = 53 * (i - 1) + 5
Excelapp.ActiveSheet.Cells(k - 3, 7) = Str(TT) + "-" + Str(i) + "ҳ" '''''ҳ
Excelapp.ActiveSheet.Cells(k - 3, 5) = QJ '''''�ڼ�
End If


Excelapp.ActiveSheet.Cells(k, 1) = Right(DT1.Recordset.Fields(0), Len(DT1.Recordset.Fields(0)) - InStr(DT1.Recordset.Fields(0), "-"))
Excelapp.ActiveSheet.Cells(k, 6) = DT1.Recordset.Fields(1)
dt2.RecordSource = "SELECT * FROM ZLCX WHERE  ��ƿ�Ŀ='" & DT1.Recordset.Fields(0) & "'"
dt2.Refresh
dt2.Recordset.MoveFirst
Do While Not dt2.Recordset.EOF
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "1" Then
Excelapp.ActiveSheet.Cells(k, 2) = Val(dt2.Recordset.Fields(7))
End If
       
If dt2.Recordset.Fields(1) = "����" Then
Excelapp.ActiveSheet.Cells(k, 4) = Val(dt2.Recordset.Fields(4))
Excelapp.ActiveSheet.Cells(k, 5) = Val(dt2.Recordset.Fields(5))
End If
       
If InStr(dt2.Recordset.Fields(1), "��") > 0 And dt2.Recordset.Fields(9) = "3" Then
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


