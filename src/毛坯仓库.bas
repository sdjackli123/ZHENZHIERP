Attribute VB_Name = "ë���ֿ�"
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
Public BDRQ As String  ''''''���ڱ������̿�¼��

Public pu As Integer
Public zu(10) As String   ''''''''''''�����ý�������
Public pd As Integer  ''''adodc6jilushu
Public khmc As String
Public zhlhji As Long
Public ww As Integer '�����Ƿ����
Public DH As String   '���ű��������ڸ���
Public Sub fhbb1(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, BT As String) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\fhbb.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�����"
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
Public Sub mprk(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ë�����.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select �ͻ�����,����,��Լ��,ë�߷���,ë��ƥ��,ë������,��ע,������,����,ny,��ɫ,����,ҵ��,��ע,������ϸ from ckgl where ���ݺ�='" & gh & "'  order by ip"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) ''�ͻ�
Excelapp.ActiveSheet.Cells(3, 5) = Trim(DT1.Recordset.Fields(8)) ''����
Excelapp.ActiveSheet.Cells(3, 8) = Trim(gh) ''���ݺ�
Excelapp.ActiveSheet.Cells(3, 11) = DT1.Recordset.Fields(7) ''����
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(12) ''˾��
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1) '''����
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(10)    ''''''''''��ɫ
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(11)    ''''''''''����
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(13)    ''''''''''��ע

' ���ñ�Ҫ��������ʽ��
Dim regexWidth As Object, regexNumber As Object
Set regexWidth = CreateObject("VBScript.RegExp")
regexWidth.Global = True
regexWidth.Pattern = "(\b\d{1,2}cm\b|��|��)" ' ƥ�����ֺ��"cm"���족���䡱

Set regexNumber = CreateObject("VBScript.RegExp")
regexNumber.Global = True
regexNumber.Pattern = "\d+(\.\d+)?" ' ƥ�����֣�����С��

' ��ȡ������ϸ
Dim details As String
details = DT1.Recordset.Fields(14) ' ���������ϸ�ڵ�15���ֶ�

' �ҵ����еķ����ǩ�����Ӧ����
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
detailArray = Split(details, " ") ' ʹ�ÿո�ָ������

' �����ָ�������
For i = LBound(detailArray) To UBound(detailArray)
    If regexWidth.Test(detailArray(i)) Then
        If currentWidth <> "" Then
            ' �洢��һ�������ƥ������������ϸ
            pieces.Add currentWidth, pieceCount
            totalWeights.Add currentWidth, weightSum
            weights.Add currentWidth, Trim(currentWeights)
            ' ���ü����������ۼ�
            pieceCount = 0
            weightSum = 0
            currentWeights = ""
        End If
        currentWidth = detailArray(i)
        ' �洢����
        widths.Add currentWidth, currentWidth
    ElseIf regexNumber.Test(detailArray(i)) Then
        ' �ۼƵ�ǰ�����ƥ��������
        pieceCount = pieceCount + 1
        weightSum = weightSum + CDbl(detailArray(i))
        currentWeights = currentWeights & detailArray(i) & " "
    End If
Next i
' �洢���һ�������ƥ������������ϸ
If currentWidth <> "" Then
    pieces.Add currentWidth, pieceCount
    totalWeights.Add currentWidth, weightSum
    weights.Add currentWidth, Trim(currentWeights)
End If

' �������Excel��Ԫ��
Dim col As Integer
col = 2 ' �ӵ�2�п�ʼ
For Each Key In widths.Keys
    Excelapp.ActiveSheet.Cells(5, col).value = widths(Key)
    Excelapp.ActiveSheet.Cells(6, col).value = pieces(Key) ' ƥ��
    Excelapp.ActiveSheet.Cells(7, col).value = totalWeights(Key) ' ������
    Excelapp.ActiveSheet.Cells(9, col).value = weights(Key) ' ������ϸ
    col = col + 2 ' ����ÿ������ռ�����У����Ը�����Ҫ����
Next Key

End If

DT1.RecordSource = "select sum(ë��ƥ��),round(sum(ë������),2) from ckgl where ���ݺ�='" & gh & "'"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub mprkf(DT1 As Adodc, gh As String, xh1 As Integer, xh2 As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ë������.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select �ͻ�����,����,��Լ��,ë�߷���,ë��ƥ��,ë������,��ע,������,����,ny from ckgl where ���ݺ�='" & gh & "' and ip between '" & xh1 & "' and '" & xh2 & "' order by ip"
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''ny���ϵ�λ
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub
Public Sub mpck(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "�ʺ��ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\mpck.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select �ͻ�,����,���,ë�߷���,ë��ƥ��,ë������,��ע,��׸���,�������� from mpbh where ����='" & gh & "'"
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
 ' ��ʾExcel�������û����б༭
    Excelapp.Visible = True
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.DisplayAlerts = False
    
    ' ������ӡԤ��
    Excelapp.ActiveSheet.PrintPreview
    
    ' ��ӡ���˳�
    Excelapp.Quit
    Set Excelapp = Nothing
    
    Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
Public Sub lcd3(DT1 As Adodc, dt2 As Adodc, gh As String, xh As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ë���뵥.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


dt2.RecordSource = "select �ͻ�����,����,����,Ʒ��,���߷���,����Ҫ��,ɫ��,ƥ��,����,��ע,��ǩ from kpd where ����='" & gh & "' and ip='" & xh & "'"
dt2.Refresh

If Not dt2.Recordset.EOF Then
dt2.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(4, 2) = dt2.Recordset.Fields(9)    '''��ע
Excelapp.ActiveSheet.Cells(6, 2) = dt2.Recordset.Fields(0)    '''�ͻ�
Excelapp.ActiveSheet.Cells(6, 6) = dt2.Recordset.Fields(3)    ''''Ʒ��
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(1)    ''''����
Excelapp.ActiveSheet.Cells(8, 4) = dt2.Recordset.Fields(2)    ''''����
Excelapp.ActiveSheet.Cells(8, 8) = dt2.Recordset.Fields(6)    ''''ɫ��
Excelapp.ActiveSheet.Cells(10, 2) = dt2.Recordset.Fields(4)    ''''����
Excelapp.ActiveSheet.Cells(10, 4) = dt2.Recordset.Fields(5)    ''''����
Excelapp.ActiveSheet.Cells(10, 6) = dt2.Recordset.Fields(7)    ''''ƥ��
Excelapp.ActiveSheet.Cells(10, 8) = dt2.Recordset.Fields(10)    ''''�ͻ�����  ���ǿ��

DT1.RecordSource = "select ����,��ע,��Ҫ,���� from sczy_x where ����='" & dt2.Recordset.Fields(1) & "' and ���='" & xh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(6, 4) = Trim(DT1.Recordset.Fields(3))    '''�ƻ�����
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1) + DT1.Recordset.Fields(2)   '''��ע��Ҫ
End If

m = 0
k = 14  '''''''''''''''''''''''''

DT1.RecordSource = "select ƥ��,���� from mpbmd where ����='" & gh & "' and ���='" & xh & "'"
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

'Excelapp.Quit '�ر�EXCEL
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



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\LCdD.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = cast('" & a & "' as real)  "
DT1.Refresh
Excelapp.ActiveSheet.Cells(2, 6) = DT1.Recordset.Fields(12)
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 12) = DT1.Recordset.Fields(8)
DT1.RecordSource = "select * from kpd where ����='" & gh & "' ORDER BY IP "
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

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub

Public Sub TMDY(TM As String, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\TMDY.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub

Public Sub ZJTMDY(TM As String, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\TMDY.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing
End Sub

Public Sub lcd2(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")
Dim lc As String


Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�Ÿ׿�ok.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & a & "'"
DT1.Refresh
Excelapp.ActiveSheet.Cells(4, 3) = DT1.Recordset.Fields(0)   ''''�ͻ�
Excelapp.ActiveSheet.Cells(2, 8) = "*" + DT1.Recordset.Fields(2) + "J*" '''����
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(2)   '''����

Excelapp.ActiveSheet.Cells(6, 3) = Trim(DT1.Recordset.Fields(12))    ''''����
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(6, 7) = "����"
Else
Excelapp.ActiveSheet.Cells(6, 7) = "����"
End If '''' ���
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''���
Excelapp.ActiveSheet.Cells(10, 3) = DT1.Recordset.Fields(52)     ''ɫ��
Excelapp.ActiveSheet.Cells(10, 7) = DT1.Recordset.Fields(8)     '''''��ɫ

Excelapp.ActiveSheet.Cells(14, 3) = DT1.Recordset.Fields(14)     ''��̨
   

''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(����,0)),2),SUM(isnull(ƥ��,0)) from kpd where ����='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(1)   ''''ƥ��
Excelapp.ActiveSheet.Cells(12, 7) = DT1.Recordset.Fields(0)    ''''�ƻ���
End If

DT1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(25 + i * 1, 2) = DT1.Recordset.Fields(55)   '''���
Excelapp.ActiveSheet.Cells(25 + i * 1, 3) = DT1.Recordset.Fields(3)   '''Ʒ��
Excelapp.ActiveSheet.Cells(25 + i * 1, 5) = DT1.Recordset.Fields(5)   '''����
Excelapp.ActiveSheet.Cells(25 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''����
Excelapp.ActiveSheet.Cells(25 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''ƥ��
Excelapp.ActiveSheet.Cells(25 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''����
Excelapp.ActiveSheet.Cells(25 + i * 1, 9) = DT1.Recordset.Fields(9)       ''''��ע ����Ҫ��
i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct ���,mr from kpd where  ����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
lc = ""
Do While Not DT1.Recordset.EOF
lc = lc + "��ţ�" + DT1.Recordset.Fields(0) + " ���̣�" + DT1.Recordset.Fields(1) + "/"  ''''''''''����
DT1.Recordset.MoveNext
Loop
End If
Excelapp.ActiveSheet.Cells(20, 3) = lc   ''''����

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
Public Sub lcd222(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�����̵�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = cast('" & a & "' as real)"
DT1.Refresh
Excelapp.ActiveSheet.Cells(5, 3) = DT1.Recordset.Fields(0)    ''''�ͻ�����
'Excelapp.ActiveSheet.Cells(4, 4) = Trim(dt1.Recordset.Fields(12))   '''''����
Excelapp.ActiveSheet.Cells(5, 17) = DT1.Recordset.Fields(2)   '''''����
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(8)  ''''����ɫ��
Excelapp.ActiveSheet.Cells(6, 10) = DT1.Recordset.Fields(52)   '''''ɫ�� �ͻ�ɫ��
Excelapp.ActiveSheet.Cells(7, 3) = DT1.Recordset.Fields(13)   ''''''��ǩ  �ͻ�����
'Excelapp.ActiveSheet.Cells(3, 2) = dt1.Recordset.Fields(3)    '''''Ʒ��
DH = DT1.Recordset.Fields(1)
xh = DT1.Recordset.Fields(11)
Excelapp.ActiveSheet.Cells(1, 16) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '��������
Excelapp.ActiveSheet.Cells(13, 3) = DT1.Recordset.Fields(30)  '''�ӹ�˵��
Excelapp.ActiveSheet.Cells(15, 3) = DT1.Recordset.Fields(51)  '''�ӹ�Ҫ��



''''''''''''''''''''''''''''
DT1.RecordSource = "select * from sczy_z where ����='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(18, 3) = DT1.Recordset.Fields(1)                ''''''''''�ܱ�ע
Excelapp.ActiveSheet.Cells(6, 3) = DT1.Recordset.Fields(0)                ''''''''''��������
End If

DT1.RecordSource = "select isnull(����,'') from sczy_x where ����='" & DH & "' and ���='" & xh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(17, 3) = DT1.Recordset.Fields(0)                ''''''''''����
End If


DT1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
i = 9
Do While Not DT1.Recordset.EOF
dt2.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
dt2.RecordSource = "select ��ע,����,���� from sczy_x where ����='" & DT1.Recordset.Fields(1) & "' and ���='" & DT1.Recordset.Fields(11) & "'"
dt2.Refresh
If Not dt2.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(i, 1) = dt2.Recordset.Fields(0)   '''֯��
Excelapp.ActiveSheet.Cells(4, 10) = Trim(dt2.Recordset.Fields(1))   '''����
Excelapp.ActiveSheet.Cells(4, 4) = Trim(dt2.Recordset.Fields(2))   '''''����
Else
Excelapp.ActiveSheet.Cells(i, 1) = ""   '''֯��
End If
Excelapp.ActiveSheet.Cells(i, 7) = DT1.Recordset.Fields(3)   '''Ʒ��
Excelapp.ActiveSheet.Cells(i, 14) = DT1.Recordset.Fields(6)   '''ƥ��
Excelapp.ActiveSheet.Cells(i, 16) = DT1.Recordset.Fields(7)   '''����
Excelapp.ActiveSheet.Cells(i, 20) = DT1.Recordset.Fields(10)  '''����
Excelapp.ActiveSheet.Cells(i, 18) = DT1.Recordset.Fields(5)  '''�ŷ�

i = i + 1
DT1.Recordset.MoveNext
Loop
End If



DT1.RecordSource = "select round(SUM(����),2),round(SUM(ƥ��),1) from kpd where ����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(6, 17) = DT1.Recordset.Fields(1)   '''�ϼ�ƥ��
Excelapp.ActiveSheet.Cells(7, 17) = DT1.Recordset.Fields(0)  ''�ϼ�����
End If



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

Public Sub jdzt(Flex As VSFlexGrid, BT As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\jdzt.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub




Public Sub jdmx(Flex As VSFlexGrid, BT As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\jdmx.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


Public Sub BBDY(Flex As VSFlexGrid, fd1, fd2, BT As String)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\bbdy.xls")
'5)���õ�2��������Ϊ�������
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
Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


Public Sub fhbb(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, BT As String) ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\fhbb.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�����"
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

'Excelapp.Quit '�ر�EXCEL
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\bhmx.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub




Public Sub sx(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''������ʾ��ʽ
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


Public Sub SX1(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''������ʾ��ʽ
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.0")
Next
End Sub

Public Sub SX2(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''������ʾ��ʽ
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.00")
Next
End Sub
Public Sub SX3(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''������ʾ��ʽ
On Error Resume Next
i = 1
p = DT.Recordset.RecordCount
For i = 1 To p
MSFlex.col = x
MSFlex.Row = i
MSFlex.Text = Format(Val(MSFlex.Text), "#0.000")
Next
End Sub
Public Sub SX4(DT As Adodc, MSFlex As VSFlexGrid, x As Integer)  '''''������ʾ��ʽ
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
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
'9) �ڵ�8��֮ǰ�����ҳ����



Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub
Public Sub MXOutadodcToExcel(Flex As VSFlexGrid, BT As String)

    Dim i   As Integer
    Dim j   As Integer
    Dim k   As Integer
    Dim Q   As Double ' ʹ��Double������ȷ���ϼ�ֵ���Դ���С��
    On Error GoTo Ert

    Dim Excelapp   As Excel.Application

    ' ����ExcelӦ�ó���ʵ��
    Set Excelapp = New Excel.Application

    On Error Resume Next

    ' �����¹������Ĺ���������
    Excelapp.SheetsInNewWorkbook = 1
    
    ' ����ExcelӦ�ó������
    Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
    
    ' ���Ѵ��ڵĹ�����
    Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
    
    ' �����һ��������
    Excelapp.Sheets(1).Activate
    
    Q = 0 ' ��ʼ���ϼ�ֵ
    
    With Flex
        k = .Rows ' ��ȡ����������

        ' ��������е�����
        For i = 1 To k
            For j = 1 To .Cols
                DoEvents
                
                ' ��鵥Ԫ���ֵ�Ƿ�Ϊ����
                If IsNumeric(.TextMatrix(i - 1, j)) Then
                    ' ��������֣�ֱ�Ӹ�ֵΪ���ָ�ʽ��ȥ��������
                    Excelapp.ActiveSheet.Cells(i + 1, j).value = CDbl(.TextMatrix(i - 1, j))
                Else
                    ' ����������֣����ı���ʽ����
                    Excelapp.ActiveSheet.Cells(i + 1, j).value = .TextMatrix(i - 1, j)
                End If
            Next j
            
            ' �ۼ�ĳ�У�FD�У�����ֵ��ȷ��FD��Ϊ��Ч�У�
            If i >= 1 Then
                Q = Q + Val(.TextMatrix(i - 1, FD))
            End If
        Next i
    End With

    ' ���ϼ�ֵ��ȷ����λС���󵼳������һ��
    Excelapp.ActiveSheet.Cells(k + 2, FD).value = Format(Q, "0.00")

    ' ����Excel��Ԫ��ı���
    Excelapp.ActiveSheet.Cells(1, 1) = BT
    
    ' ���ô������ű���
    Excelapp.ActiveWindow.Zoom = 100
    
    ' ��ʾExcelӦ�ó���
    Excelapp.Visible = True
    
    ' ���þ�����ʾ
    Excelapp.DisplayAlerts = False
    
    ' �˳������ExcelӦ�ó���ʵ��
    Set Excelapp = Nothing
    Excelapp.Quit
    Exit Sub

Ert:
    ' �������˳�Excel
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

        
Excelapp.Caption = "����Ⱦ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\��Լ��.xls")
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

'Excelapp.Quit '�ر�EXCEL
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\����.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
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

        
Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ӡȾ.xls")
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

Public Sub YEBDOutadodcToExcel(DT1 As Adodc, dt2 As Adodc, QJ As String)  ''''�ޱ���

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
Excelapp.ActiveSheet.Cells(k, 4) = dt2.Recordset.Fields(4)
Excelapp.ActiveSheet.Cells(k, 5) = dt2.Recordset.Fields(5)
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
MsgBox ("")
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintPreview
        Set Excelapp = Nothing
        Excelapp.Quit
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

Public Sub OutadodcToExcel3(Flex As VSFlexGrid, fd1, fd2, fd3, BT) ''''��һ�ֶκϼƣ������⣩

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "������ñ���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub
Public Sub OutadodcToExcel22(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, fd5, fd6, fd7, fd8, fd9, fd10, fd11, fd12, fd13, BT) ''''��һ�ֶκϼƣ������⣩

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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub
Public Sub OutadodcToExcel2(Flex As VSFlexGrid, fd1, fd2, BT) ''''��һ�ֶκϼƣ������⣩

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

'Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub

Public Sub OutadodcToExcel4(Flex As VSFlexGrid, fd1, fd2, fd3, fd4, BT) ''''��һ�ֶκϼƣ������⣩

        Dim i   As Integer
        Dim j   As Integer
        Dim k   As Integer
        Dim x   As Integer
        On Error GoTo Ert



        Dim Excelapp   As Excel.Application

        Set Excelapp = New Excel.Application

        On Error Resume Next


        Excelapp.SheetsInNewWorkbook = 1

        
Excelapp.Caption = "������ñ���֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\lbj.xls")
'5)���õ�2��������Ϊ�������
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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub



Public Sub lcd33(DT1 As Adodc, dt2 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�Ÿ׿�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select distinct ����,�ͻ�����,CONVERT(varchar,����, 23),��ǩ from kpd where ����='" & gh & "' order by ����"
DT1.Refresh

If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)
Excelapp.ActiveSheet.Cells(2, 7) = Trim(DT1.Recordset.Fields(2))
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(2, 13) = Trim(gh)

i = 4
Do While Not DT1.Recordset.EOF

dt2.RecordSource = "select * from kpd where ����='" & DT1.Recordset.Fields(0) & "' order by IP"
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

'Excelapp.Quit '�ر�EXCEL
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



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���̵�ok.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & a & "'"
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




''''''''''''''''''''''''''''С��
'Excelapp.ActiveSheet.Cells(22, 2) = dt1.Recordset.Fields(0)
'Excelapp.ActiveSheet.Cells(22, 5) = dt1.Recordset.Fields(14)
'Excelapp.ActiveSheet.Cells(22, 7) = dt1.Recordset.Fields(15)
'Excelapp.ActiveSheet.Cells(23, 1) = dt1.Recordset.Fields(8)
'Excelapp.ActiveSheet.Cells(23, 5) = dt1.Recordset.Fields(2)

'Excelapp.ActiveSheet.Cells(27, 2) = dt1.Recordset.Fields(3)
'Excelapp.ActiveSheet.Cells(31, 2) = dt1.Recordset.Fields(9) + Space(5) + "����:" + dt1.Recordset.Fields(5) + "   ����" + dt1.Recordset.Fields(10)
Excelapp.ActiveSheet.Cells(5, 9) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '��������

DT1.RecordSource = "select * from sczy_z where ����='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)                ''''''''''�ܱ�ע
End If


'dt1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
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

DT1.RecordSource = "select round(SUM(����),2),round(SUM(ƥ��),1) from kpd where ����='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
End If

DT1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
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
Excelapp.ActiveSheet.Cells(10 + i * 2, 9) = DT1.Recordset.Fields(30)     ''''''�ӹ���Ŀ
Excelapp.ActiveSheet.Cells(10 + i * 2, 12) = DT1.Recordset.Fields(9)     '''��ע
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

'Excelapp.Quit '�ر�EXCEL
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



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���̵�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "select ����,����,Ʒ��,���߷���,����,ɫ��,��ע,����,ͶȾ���,������;,��ͬ����,�µ�����,��ͬ����,�ܱ�ע,�ƻ�����,�ɷ�,����Ҫ��,����,ɫ��,��������,����,��ǩ,ƥ�� from v_kpd_mx where ����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(1, 9) = "*" + DT1.Recordset.Fields(1) + "J*" ''''��������
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(8)   ''''�ӹ�����
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(10)  '''ҵ��
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(7))   '''����
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(19)   '''��������
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(15)   ''''�ɷ�
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(9)    '''������;
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(0)   '''��ͬ��
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)   '''����
Excelapp.ActiveSheet.Cells(5, 9) = DT1.Recordset.Fields(5)  '��ɫ
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(1)   '''�׺�
Excelapp.ActiveSheet.Cells(6, 6) = Trim(DT1.Recordset.Fields(12))    '''����
Excelapp.ActiveSheet.Cells(6, 9) = DT1.Recordset.Fields(18)   'ɫ��
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)   '''����
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(21)   ''ԭ���� �ȿ��
If Val(DT1.Recordset.Fields(22)) = 0 Then
Excelapp.ActiveSheet.Cells(7, 9) = "" '�ƻ�ƥ��
Else
Excelapp.ActiveSheet.Cells(7, 9) = DT1.Recordset.Fields(22)  '�ƻ�ƥ��
End If
Excelapp.ActiveSheet.Cells(7, 12) = ""  'ʵ��ƥ��
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(3)   '''����
Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(16)    '''����
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(4)  '�ƻ�����
Excelapp.ActiveSheet.Cells(8, 12) = "" 'ʵ������
Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(20)   ''''����
Excelapp.ActiveSheet.Cells(13, 2) = DT1.Recordset.Fields(13)   ''''�ܱ�ע
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(6)   ''''��ע
End If
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
Public Sub lcd22yh(DT1 As Adodc, gh As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")



'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ӡ�����̵�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.RecordSource = "select ����,����,Ʒ��,���߷���,����,ɫ��,��ע,����,ͶȾ���,������;,��ͬ����,�µ�����,��ͬ����,�ܱ�ע,�ƻ�����,�ɷ�,����Ҫ��,����,ɫ��,��������,����,��ǩ,ƥ�� from v_kpd_mx where ����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(1, 9) = "*" + DT1.Recordset.Fields(1) + "J*" ''''��������
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(8)   ''''�ӹ�����
Excelapp.ActiveSheet.Cells(3, 6) = DT1.Recordset.Fields(10)  '''ҵ��
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(7))   '''����
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(19)   '''��������
Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(15)   ''''�ɷ�
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(9)    '''������;
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(0)   '''��ͬ��
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(17)   '''����
Excelapp.ActiveSheet.Cells(5, 9) = DT1.Recordset.Fields(5)  '��ɫ
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(1)   '''�׺�
Excelapp.ActiveSheet.Cells(6, 6) = Trim(DT1.Recordset.Fields(12))    '''����
Excelapp.ActiveSheet.Cells(6, 9) = DT1.Recordset.Fields(18)   'ɫ��
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)   '''����
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(21)   ''ԭ���� �ȿ��
If Val(DT1.Recordset.Fields(22)) = 0 Then
Excelapp.ActiveSheet.Cells(7, 9) = ""  '�ƻ�ƥ��
Else
Excelapp.ActiveSheet.Cells(7, 9) = DT1.Recordset.Fields(22)  '�ƻ�ƥ��
End If
Excelapp.ActiveSheet.Cells(7, 12) = ""  'ʵ��ƥ��
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(3)   '''����
Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(16)    '''����
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(4)  '�ƻ�����
Excelapp.ActiveSheet.Cells(8, 12) = "" 'ʵ������
Excelapp.ActiveSheet.Cells(10, 1) = DT1.Recordset.Fields(20)   ''''����
Excelapp.ActiveSheet.Cells(13, 2) = DT1.Recordset.Fields(13)   ''''�ܱ�ע
Excelapp.ActiveSheet.Cells(15, 2) = DT1.Recordset.Fields(6)   ''''��ע
End If
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
Public Sub lcd22fx(DT1 As Adodc, gh As String, lb As String)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application
Dim b As Integer
Dim DH As String
On Error Resume Next

'adodcEnvironment1.kp Text7.Text, b
'adodcReport1.Show 1
'adodcEnvironment1.rskp.Close



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���̵���.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")
DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & a & "' and ����='" & lb & "'"
DT1.Refresh

b = DT1.Recordset.Fields(11)   ''''���
DH = DT1.Recordset.Fields(1)   ''''����

Excelapp.ActiveSheet.Cells(1, 2) = DT1.Recordset.Fields(51)    ''����  '''��֯��Ҫ��
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)    ''�ͻ�
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(14)   ''��̨
Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(3)    ''Ʒ��
Excelapp.ActiveSheet.Cells(3, 13) = DT1.Recordset.Fields(8) + DT1.Recordset.Fields(52) ''ɫ��   ��ɫ+ɫ��
Excelapp.ActiveSheet.Cells(2, 10) = Trim(DT1.Recordset.Fields(12))   ''''����
Excelapp.ActiveSheet.Cells(2, 13) = DT1.Recordset.Fields(2)        '''''����
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(13)        ''''���
Excelapp.ActiveSheet.Cells(10, 3) = DT1.Recordset.Fields(5)        ''''��������
Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(10)        ''''����
Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(9)        ''''��ע    ȾɫҪ��  �ܱ�ע
Excelapp.ActiveSheet.Cells(3, 9) = Trim(DT1.Recordset.Fields(6))       ''''�ƻ�ƥ
Excelapp.ActiveSheet.Cells(4, 9) = Trim(DT1.Recordset.Fields(7))        ''''�ƻ���
Excelapp.ActiveSheet.Cells(5, 9) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '��������
Excelapp.ActiveSheet.Cells(12, 13) = DT1.Recordset.Fields(46)    ''  '''ӡ��ͼ��


''''''''''''''''''''''''''''

DT1.RecordSource = "select ��ע,���� as �ܱ�ע,���� as ����,�ɷ� as ����,��ˮ�� as ��ˮ,Ť�� as �ָ�,���� as ˵�� from sczy_x where ����='" & DH & "' and ���='" & b & "'"   '''���� ֯���� �ɷ�--���� ��������
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 6) = DT1.Recordset.Fields(2)                ''''''''''֯��
Excelapp.ActiveSheet.Cells(7, 6) = DT1.Recordset.Fields(3)                ''''''''''����
Excelapp.ActiveSheet.Cells(10, 5) = DT1.Recordset.Fields(0)                ''''''''''��ע
'Excelapp.ActiveSheet.Cells(12, 3) = DT1.Recordset.Fields(1)                ''''''''''�ܱ�ע
Excelapp.ActiveSheet.Cells(31, 2) = DT1.Recordset.Fields(4)                ''''''''''��ˮ��
Excelapp.ActiveSheet.Cells(32, 2) = DT1.Recordset.Fields(5)                ''''''''''��ˮ��
Excelapp.ActiveSheet.Cells(31, 8) = DT1.Recordset.Fields(6)                ''''''''''˵��
End If

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


Public Sub mprkbqdy(DT1 As Adodc, gh As String, xh As Integer)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next

Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\mprkbq.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select �ͻ�����,����,��Լ��,ë�߷���,ë��ƥ��,ë������,��ע,������,����,ny,���λ�� from ckgl where ���ݺ�='" & gh & "' and ip='" & xh & "'"
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

'Excelapp.Quit '�ر�EXCEL
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



Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ί�����̵���ok.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = cast('" & a & "' as real)"
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
'Excelapp.ActiveSheet.Cells(31, 2) = dt1.Recordset.Fields(9) + Space(5) + "����:" + dt1.Recordset.Fields(5) + "   ����" + dt1.Recordset.Fields(10)
'Excelapp.ActiveSheet.Cells(5, 9) = "*" + dt1.Recordset.Fields(2) + "J" + "*"  '��������

DT1.RecordSource = "select * from sczy_z where ����='" & DH & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)                ''''''''''�ܱ�ע
End If

DT1.RecordSource = "select * from kpdwwjg where ί�����='" & gh & "'"
DT1.Refresh
If Not DT1.Recordset.EOF Then
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(1)                ''''''''''ί�ⵥλ
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(2)                ''''''''''ί����Ϣ
End If


'dt1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
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

DT1.RecordSource = "select round(SUM(����),2),round(SUM(ƥ��),1) from kpd where ����='" & gh & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 9) = DT1.Recordset.Fields(0)
Excelapp.ActiveSheet.Cells(3, 9) = DT1.Recordset.Fields(1)
End If

DT1.RecordSource = "select * from kpd where ����='" & gh & "' order by IP"
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
Excelapp.ActiveSheet.Cells(10 + i * 2, 9) = Trim(DT1.Recordset.Fields(30)) ''''''�ӹ���Ŀ
Excelapp.ActiveSheet.Cells(10 + i * 2, 12) = DT1.Recordset.Fields(9)  '''''+��ע
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
    Dim Excelapp As Object ' ʹ��ͨ�ö�������
    Dim Workbook As Object
    Dim Worksheet As Object
    
    On Error GoTo Ert

    ' ���� Excel Ӧ�ó������
    Set Excelapp = CreateObject("Excel.Application")
    If Excelapp Is Nothing Then
        MsgBox "�޷����� Excel Ӧ�ó��������ȷ���Ѱ�װ Excel�������Ȩ�޺�ע������á�"
        Exit Sub
    End If

    ' ���� Excel Ӧ�ó�������
    Excelapp.Caption = "������ӡģ�����֮��ӡ"
    Excelapp.SheetsInNewWorkbook = 1

    ' �����еĹ�����
    Set Workbook = Excelapp.Workbooks.Open(App.Path & "\��ӡģ��\����\�������̿�.xls")
    Set Worksheet = Workbook.Sheets(1)
    Worksheet.Activate

    ' �������ݿ����ӺͲ�ѯ
    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "' and ����='" & lb & "'"
    DT1.Refresh
    If DT1.Recordset.EOF Then GoTo Cleanup

    Dim maxWeight As Double
    maxWeight = DT1.Recordset.Fields("zl")

    DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & maxWeight & "' and ����='" & lb & "'"
    DT1.Refresh
    If DT1.Recordset.EOF Then GoTo Cleanup


'Excelapp.ActiveSheet.Cells(3, 3) = Trim(lb)   ''''����
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(12) '''����������
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(2) '''����
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''�ͻ�
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(3) '''����
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8) '''��ɫ
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(52) '''ɫ��
Excelapp.ActiveSheet.Cells(9, 2) = DT1.Recordset.Fields(9) '''ȾɫҪ��

Excelapp.ActiveSheet.Cells(2, 8) = DT1.Recordset.Fields(0) '''�Ÿ׿��ͻ�
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 10) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(52) ''ɫ��
Excelapp.ActiveSheet.Cells(5, 8) = DT1.Recordset.Fields(5) ''����
Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(10) ''����
Excelapp.ActiveSheet.Cells(7, 7) = DT1.Recordset.Fields(9) '''ȾɫҪ��
Excelapp.ActiveSheet.Cells(11, 7) = DT1.Recordset.Fields(12) '''�Ÿ׿�����


Excelapp.ActiveSheet.Cells(21, 2) = DT1.Recordset.Fields(0)   ''''�ͻ�
Excelapp.ActiveSheet.Cells(18, 6) = "*" + DT1.Recordset.Fields(2) + "J*" '''����
Excelapp.ActiveSheet.Cells(20, 5) = DT1.Recordset.Fields(2)   '''����

Excelapp.ActiveSheet.Cells(20, 2) = Trim(DT1.Recordset.Fields(12))    ''''����
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(18, 1) = "����"
Else
Excelapp.ActiveSheet.Cells(18, 1) = "����"
End If '''' ���
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''���
Excelapp.ActiveSheet.Cells(23, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''ɫ��+��ɫ
Excelapp.ActiveSheet.Cells(24, 2) = DT1.Recordset.Fields(9)     '''''ȾɫҪ��
Excelapp.ActiveSheet.Cells(22, 2) = DT1.Recordset.Fields(3)   ''Ʒ��
Excelapp.ActiveSheet.Cells(26, 2) = DT1.Recordset.Fields(5)     ''����
 Excelapp.ActiveSheet.Cells(23, 6) = DT1.Recordset.Fields(82) ''������ϸ
Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(10)  ''����
''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(����,0)),2),SUM(isnull(ƥ��,0)) from kpd where ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(1)   ''''������ƥ��
Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(0)    ''''�������ƻ���
Excelapp.ActiveSheet.Cells(4, 10) = DT1.Recordset.Fields(1)   ''''�Ÿ׿�ƥ��
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(0)    ''''�Ÿ׿��ƻ���

Excelapp.ActiveSheet.Cells(28, 2) = DT1.Recordset.Fields(1)   ''''ƥ��
Excelapp.ActiveSheet.Cells(29, 2) = DT1.Recordset.Fields(0)    ''''�ƻ���
End If

DT1.RecordSource = "select * from kpd where ����='" & gh & "'  order by ����,IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(16 + i * 1, 2) = DT1.Recordset.Fields(55)   '''���
Excelapp.ActiveSheet.Cells(16 + i * 1, 3) = DT1.Recordset.Fields(3)   '''Ʒ��
Excelapp.ActiveSheet.Cells(16 + i * 1, 5) = DT1.Recordset.Fields(5)   '''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''ƥ��
Excelapp.ActiveSheet.Cells(16 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 9) = DT1.Recordset.Fields("����")       ''''����

i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct ���,mr from kpd where  ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''����
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(31, 2) = lc  ''''����

'''�����齫���̷ֿ������Ŵ�ӡ�ڱ����
'Dim dataArray() As String
'dataArray = Split(lc, "-")

'Dim L As Integer
'For L = 0 To UBound(dataArray)
'   Excelapp.ActiveSheet.Cells(L + 38, 1).value = dataArray(L)
'Next L


'DT1.RecordSource = "select distinct ���,��ע from kpd where  ����='" & gh & "' and ����='" & lb & "'"
'DT1.Refresh
'If Not DT1.Recordset.EOF Then
'bz = ""
'xbz = ""
'Do While Not DT1.Recordset.EOF

'If InStr(xbz, DT1.Recordset.Fields(1)) = 0 Then
'xbz = xbz + DT1.Recordset.Fields(1)
'End If

dt2.RecordSource = "select * from ckgl where ���ݺ�='" & gh & "'"     ''���ﵥ�ݺű������gh���ܵ���ҵ��
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If
Excelapp.ActiveSheet.Cells(5, 2) = dt2.Recordset.Fields(12) ''���������ϵ�λ
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(9) '''���������λ��

Excelapp.ActiveSheet.Cells(20, 8) = dt2.Recordset.Fields(16)  ''˾��ҵ��
Excelapp.ActiveSheet.Cells(21, 8) = dt2.Recordset.Fields(12) ''���ϵ�λ
Excelapp.ActiveSheet.Cells(30, 2) = dt2.Recordset.Fields(9) '''���λ��
End If


    ' ���ô�ӡ���Բ���ӡ
    Excelapp.ActiveWindow.Zoom = 100
    Excelapp.DisplayAlerts = False

    ' �л����û�ѡ��Ĵ�ӡ��
    If selectedPrinter <> "" Then
        TrySetActivePrinter Excelapp, selectedPrinter
    End If

    ' ��ӡ������
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
        ' ���Ը��Ӷ˿�����
        Excelapp.ActivePrinter = PrinterName & " on " & Split(PrinterName, " (")(1) ' ��ȡ�����Ӷ˿�����
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


Excelapp.Caption = "����Ⱦ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�������̿�.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & a & "' and ����='" & lb & "'"
DT1.Refresh
'Excelapp.ActiveSheet.Cells(3, 3) = Trim(lb)   ''''����
Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(12) '''����������
Excelapp.ActiveSheet.Cells(2, 5) = DT1.Recordset.Fields(2) '''����
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''�ͻ�
Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(3) '''����
Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8) '''��ɫ
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(52) '''ɫ��
Excelapp.ActiveSheet.Cells(9, 2) = DT1.Recordset.Fields(9) '''ȾɫҪ��

Excelapp.ActiveSheet.Cells(2, 8) = DT1.Recordset.Fields(0) '''�Ÿ׿��ͻ�
Excelapp.ActiveSheet.Cells(2, 10) = DT1.Recordset.Fields(2)
Excelapp.ActiveSheet.Cells(3, 8) = DT1.Recordset.Fields(3)
Excelapp.ActiveSheet.Cells(3, 10) = DT1.Recordset.Fields(8)
Excelapp.ActiveSheet.Cells(4, 8) = DT1.Recordset.Fields(52) ''ɫ��
Excelapp.ActiveSheet.Cells(5, 8) = DT1.Recordset.Fields(5) ''����
Excelapp.ActiveSheet.Cells(6, 8) = DT1.Recordset.Fields(10) ''����
Excelapp.ActiveSheet.Cells(7, 7) = DT1.Recordset.Fields(9) '''ȾɫҪ��
Excelapp.ActiveSheet.Cells(11, 7) = DT1.Recordset.Fields(12) '''�Ÿ׿�����


Excelapp.ActiveSheet.Cells(21, 2) = DT1.Recordset.Fields(0)   ''''�ͻ�
Excelapp.ActiveSheet.Cells(18, 6) = "*" + DT1.Recordset.Fields(2) + "J*" '''����
Excelapp.ActiveSheet.Cells(20, 5) = DT1.Recordset.Fields(2)   '''����

Excelapp.ActiveSheet.Cells(20, 2) = Trim(DT1.Recordset.Fields(12))    ''''����
If InStr(DT1.Recordset.Fields(2), "F") > 0 Or InStr(DT1.Recordset.Fields(2), "H") > 0 Then
Excelapp.ActiveSheet.Cells(18, 1) = "����"
Else
Excelapp.ActiveSheet.Cells(18, 1) = "����"
End If '''' ���
Excelapp.ActiveSheet.Cells(8, 7) = DT1.Recordset.Fields(13)     ''''���
Excelapp.ActiveSheet.Cells(23, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''ɫ��+��ɫ
Excelapp.ActiveSheet.Cells(24, 2) = DT1.Recordset.Fields(9)     '''''ȾɫҪ��
Excelapp.ActiveSheet.Cells(22, 2) = DT1.Recordset.Fields(3)   ''Ʒ��
Excelapp.ActiveSheet.Cells(26, 2) = DT1.Recordset.Fields(5)     ''����
 Excelapp.ActiveSheet.Cells(23, 6) = DT1.Recordset.Fields(82) ''������ϸ
Excelapp.ActiveSheet.Cells(27, 2) = DT1.Recordset.Fields(10)  ''����
''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(����,0)),2),SUM(isnull(ƥ��,0)) from kpd where ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(1)   ''''������ƥ��
Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(0)    ''''�������ƻ���
Excelapp.ActiveSheet.Cells(4, 10) = DT1.Recordset.Fields(1)   ''''�Ÿ׿�ƥ��
Excelapp.ActiveSheet.Cells(5, 10) = DT1.Recordset.Fields(0)    ''''�Ÿ׿��ƻ���

Excelapp.ActiveSheet.Cells(28, 2) = DT1.Recordset.Fields(1)   ''''ƥ��
Excelapp.ActiveSheet.Cells(29, 2) = DT1.Recordset.Fields(0)    ''''�ƻ���
End If

DT1.RecordSource = "select * from kpd where ����='" & gh & "'  order by ����,IP"
DT1.Refresh
i = 0
DT1.Recordset.MoveFirst
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(16 + i * 1, 2) = DT1.Recordset.Fields(55)   '''���
Excelapp.ActiveSheet.Cells(16 + i * 1, 3) = DT1.Recordset.Fields(3)   '''Ʒ��
Excelapp.ActiveSheet.Cells(16 + i * 1, 5) = DT1.Recordset.Fields(5)   '''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 6) = Trim(DT1.Recordset.Fields(10))  '''''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 7) = Trim(DT1.Recordset.Fields(6))  ''''ƥ��
Excelapp.ActiveSheet.Cells(16 + i * 1, 8) = Trim(DT1.Recordset.Fields(7))         ''''����
Excelapp.ActiveSheet.Cells(16 + i * 1, 9) = DT1.Recordset.Fields("����")       ''''����

i = i + 1
DT1.Recordset.MoveNext
Loop

DT1.RecordSource = "select distinct ���,mr from kpd where  ����='" & gh & "' and ����='" & lb & "'"
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''����
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(31, 2) = lc  ''''����

'''�����齫���̷ֿ������Ŵ�ӡ�ڱ����
'Dim dataArray() As String
'dataArray = Split(lc, "-")

'Dim L As Integer
'For L = 0 To UBound(dataArray)
'   Excelapp.ActiveSheet.Cells(L + 38, 1).value = dataArray(L)
'Next L


'DT1.RecordSource = "select distinct ���,��ע from kpd where  ����='" & gh & "' and ����='" & lb & "'"
'DT1.Refresh
'If Not DT1.Recordset.EOF Then
'bz = ""
'xbz = ""
'Do While Not DT1.Recordset.EOF

'If InStr(xbz, DT1.Recordset.Fields(1)) = 0 Then
'xbz = xbz + DT1.Recordset.Fields(1)
'End If

dt2.RecordSource = "select * from ckgl where ���ݺ�='" & gh & "'"     ''���ﵥ�ݺű������gh���ܵ���ҵ��
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If
Excelapp.ActiveSheet.Cells(5, 2) = dt2.Recordset.Fields(12) ''���������ϵ�λ
Excelapp.ActiveSheet.Cells(8, 2) = dt2.Recordset.Fields(9) '''���������λ��

Excelapp.ActiveSheet.Cells(20, 8) = dt2.Recordset.Fields(16)  ''˾��ҵ��
Excelapp.ActiveSheet.Cells(21, 8) = dt2.Recordset.Fields(12) ''���ϵ�λ
Excelapp.ActiveSheet.Cells(30, 2) = dt2.Recordset.Fields(9) '''���λ��
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
Public Sub lcd22f2(DT1 As Adodc, dt2 As Adodc, gh As String, lb As String)
    Dim Excelapp As Excel.Application
    Set Excelapp = New Excel.Application
    Excelapp.Visible = False  ' Initially hide the application to prevent screen flickering

    On Error GoTo errorhandler

    ' Open the Excel template
    Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\���̿�ok.xls")
    Excelapp.Sheets(1).Activate
    DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
    DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "' and ����='" & lb & "'"
    DT1.Refresh
    Dim maxWeight As Variant
    maxWeight = DT1.Recordset.Fields("zl").value

    DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & maxWeight & "' and ����='" & lb & "'"
    DT1.Refresh

    With Excelapp.ActiveSheet
        .Cells(3, 9).value = DT1.Recordset.Fields(2).value ' ����
        .Cells(3, 2).value = DT1.Recordset.Fields(0).value ' �ͻ�
        .Cells(4, 2).value = DT1.Recordset.Fields(3).value ' ����
        .Cells(3, 6).value = DT1.Recordset.Fields(8).value ' ��ɫ
        .Cells(4, 6).value = DT1.Recordset.Fields(52).value ' ɫ��
        .Cells(6, 3).value = DT1.Recordset.Fields(9).value ' ȾɫҪ��
        .Cells(5, 2).value = DT1.Recordset.Fields(13).value ' ���
        .Cells(8, 2).value = DT1.Recordset.Fields(5).value ' ����
        .Cells(8, 6).value = DT1.Recordset.Fields(10).value ' ����
        .Cells(1, 7) = "*" + DT1.Recordset.Fields(2) + "J" + "*"  '��������
  ' ��ȡ������ϸ����
    Dim widthDetails As String
    widthDetails = DT1.Recordset.Fields("������ϸ").value
    Dim items() As String
    items = Split(widthDetails, " ")
    
    ' ������ʼ�м���
    Dim startRow As Integer, column As Integer
    startRow = 11    ' �ӵ�11�п�ʼ
    column = 2      ' �ӵ�2�п�ʼ

    ' ���嵱ǰ���Ѵ�ӡ������
    Dim printedRows As Integer
    printedRows = 0

    ' ����������
    For i = LBound(items) To UBound(items)
        ' �ж��Ƿ��������+cm��������Ӵ���ʾ�������������СΪ18��
        If InStr(items(i), "cm") > 0 Then
            .Cells(startRow + printedRows, column).value = items(i)
            .Cells(startRow + printedRows, column).Font.Bold = True ' �Ӵ���ʾ
            .Cells(startRow + printedRows, column).Font.Size = 18 ' ���������СΪ18��
        ElseIf InStr(items(i), "��") > 0 Or InStr(items(i), "��") > 0 Then
            .Cells(startRow + printedRows, column).value = items(i)
            .Cells(startRow + printedRows, column).Font.Bold = True ' �Ӵ���ʾ
            .Cells(startRow + printedRows, column).Font.Size = 18 ' ���������СΪ18��
        Else
            .Cells(startRow + printedRows, column).value = items(i)
            ' ��Բ�������+cm�������ȡ���Ӵ���ʾ�������������СΪ18��
            .Cells(startRow + printedRows, column).Font.Bold = False
            .Cells(startRow + printedRows, column).Font.Size = 18 ' ���������СΪ18��
        End If
        
        .Cells(startRow + printedRows, column).WrapText = False  ' ���û���

        ' ����ӡ����������26ʱ���л�����һ�в����ô�ӡ����
        printedRows = printedRows + 1
        If printedRows >= 26 Then
            column = column + 1    ' �л�����һ��
            printedRows = 0  ' ���ô�ӡ����
        End If
    Next i
End With

DT1.RecordSource = "select round(SUM(isnull(����,0)),2),SUM(isnull(ƥ��,0)) from kpd where ����='" & gh & "' "
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else
Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(1) & "ƥ" ' ��ƥ��
Excelapp.ActiveSheet.Cells(8, 9) = DT1.Recordset.Fields(0) & "kg" ' ������
End If
dt2.RecordSource = "select * from ckgl where ���ݺ�='" & gh & "'"     ''���ﵥ�ݺű������gh���ܵ���ҵ��
dt2.Refresh
If Not dt2.Recordset.EOF Then

Excelapp.ActiveSheet.Cells(5, 9) = dt2.Recordset.Fields(16)  ''˾��ҵ��
Excelapp.ActiveSheet.Cells(4, 9) = dt2.Recordset.Fields(12) ''���ϵ�λ

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


Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\ë�����.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate


DT1.RecordSource = "select �׺�,����,sum(ë��ƥ��),round(sum(ë������),2),�������� from mpbh where ����='" & gh & "' and �׺� in(select distinct ��� from kpd where ����='" & gh & "' and ����='" & kh & "') group by �׺�,����,�������� order by �׺�"
DT1.Refresh
L = 0
If Not DT1.Recordset.EOF Then
DT1.Recordset.MoveFirst
Excelapp.ActiveSheet.Cells(31, 6) = Trim(DT1.Recordset.Fields(4))   ''''��������
Do While Not DT1.Recordset.EOF
Excelapp.ActiveSheet.Cells(33 + L, 2) = DT1.Recordset.Fields(0)   ''''�׺�
Excelapp.ActiveSheet.Cells(33 + L, 4) = DT1.Recordset.Fields(1)   ''����
Excelapp.ActiveSheet.Cells(33 + L, 6) = Trim(DT1.Recordset.Fields(3))   '''''����
Excelapp.ActiveSheet.Cells(33 + L, 7) = Trim(DT1.Recordset.Fields(2))   '''''ƥ��
L = L + 1
DT1.Recordset.MoveNext
Loop
End If
If kh = "��" Then
Excelapp.ActiveSheet.Cells(38, 2) = "����"  '''''�ϼ�ƥ��
Else
Excelapp.ActiveSheet.Cells(38, 2) = "����"  '''''�ϼ�ƥ��
End If

DT1.RecordSource = "select isnull(sum(ë��ƥ��),0),isnull(round(sum(ë������),2),0) from mpbh where ����='" & gh & "'"
DT1.Refresh
Excelapp.ActiveSheet.Cells(38, 2) = "�ϼ�"  '''''�ϼ�ƥ��
Excelapp.ActiveSheet.Cells(38, 7) = Trim(DT1.Recordset.Fields(0))   '''''�ϼ�ƥ��
Excelapp.ActiveSheet.Cells(38, 6) = Trim(DT1.Recordset.Fields(1))   '''''�ϼ�����

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

    ' ������ʽƥ��������ֺ��"cm"�ͺ�������֣��򵥶���"��"��"��"�ͺ��������
    regex.Pattern = "(\d+\s*cm\s*(\d+\s*)+)|(��\s*(\d+\s*)+)|(��\s*(\d+\s*)+)"
    regex.Global = True
    regex.IgnoreCase = True

    ' ִ��ƥ��
    Set matches = regex.Execute(Data)

    Dim match As match
    For Each match In matches
        ' ����һ���µļ������������Ͷ�Ӧ������
        Dim itemCollection As New Collection
        Dim content As String
        content = match.value

        ' ʹ�ÿո�ָ���������
        Dim parts() As String
        parts = Split(content, " ")

        Dim i As Integer
        For i = 1 To UBound(parts)
            If IsNumeric(parts(i)) Then
                itemCollection.Add parts(i) ' �������������
            End If
        Next

        widthDetails.Add itemCollection, parts(0) ' ʹ�÷�����Ϊ��
    Next

    Set ExtractWidthAndWeights = widthDetails
End Function

Public Sub lcd22f3(DT1 As Adodc, dt2 As Adodc, gh As String, weight As Double, count As Double)
Dim Excelapp   As Excel.Application
Set Excelapp = New Excel.Application

On Error Resume Next
b = DT1.Recordset.Fields("ip")
Dim lc As String


Excelapp.Caption = "���˴�ӡģ�����֮��ӡ"
'3)����¹�������
'4)���Ѵ��ڵĹ�������
Excelapp.SheetsInNewWorkbook = 1
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\�������.xls")
'5)���õ�2��������Ϊ�������
Excelapp.Sheets(1).Activate

DT1.ConnectionString = "Provider=SQLOLEDB.1;Password=fjrj;Persist Security Info=True;User ID=sa;Initial Catalog=zzpr;Data Source=192.168.1.254"
DT1.RecordSource = "select max(����) as zl from kpd where ����='" & gh & "' "
DT1.Refresh
a = DT1.Recordset.Fields("zl")

DT1.RecordSource = "select * from kpd where ����='" & gh & "' And ���� = '" & a & "' "
DT1.Refresh


Excelapp.ActiveSheet.Cells(2, 2) = DT1.Recordset.Fields(0)   ''''�ͻ�
Excelapp.ActiveSheet.Cells(1, 5) = DT1.Recordset.Fields(2)   '''����

Excelapp.ActiveSheet.Cells(1, 2) = Trim(DT1.Recordset.Fields(12))    ''''����

Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(52) + DT1.Recordset.Fields(8)   ''ɫ��+��ɫ
Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(9)     '''''ȾɫҪ��
Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(3)   ''Ʒ��
Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(5)     ''����
 Excelapp.ActiveSheet.Cells(4, 6) = DT1.Recordset.Fields(82) ''������ϸ
Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(10)  ''����
Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(83) ''��ͷ


''''''''''''''''''''''''''''
DT1.RecordSource = "select round(SUM(isnull(����,0)),2),SUM(isnull(ƥ��,0)) from kpd where ����='" & gh & "' "
DT1.Refresh
If DT1.Recordset.EOF Then
        Set Excelapp = Nothing
        Excelapp.Quit
Exit Sub
Else

Excelapp.ActiveSheet.Cells(9, 2).value = count   ' ƥ��
Excelapp.ActiveSheet.Cells(10, 2).value = weight ' �ƻ���

End If
DT1.RecordSource = "select distinct ���,mr from kpd where  ����='" & gh & "' "
DT1.Refresh

If Not DT1.Recordset.EOF Then
    DT1.Recordset.MoveFirst
    lc = ""
    Do While Not DT1.Recordset.EOF
        If InStr(lc, DT1.Recordset.Fields(1)) = 0 Then
            lc = lc + DT1.Recordset.Fields(1)   ''''''''''����
        End If
        DT1.Recordset.MoveNext
    Loop
End If
Excelapp.ActiveSheet.Cells(12, 2) = lc  ''''����


dt2.RecordSource = "select * from ckgl where ���ݺ�='" & gh & "'"     ''���ﵥ�ݺű������gh���ܵ���ҵ��
dt2.Refresh
If Not dt2.Recordset.EOF Then
'If InStr(bz, dt2.Recordset.Fields(0)) = 0 Then
'bz = bz + dt2.Recordset.Fields(0)
'End If
'End If
'DT1.Recordset.MoveNext
'Loop
'End If

Excelapp.ActiveSheet.Cells(8, 5) = dt2.Recordset.Fields(18) ''������
Excelapp.ActiveSheet.Cells(9, 5) = dt2.Recordset.Fields(21) ''��ƥ��
Excelapp.ActiveSheet.Cells(10, 5) = dt2.Recordset.Fields(20) ''����ƥ��
Excelapp.ActiveSheet.Cells(11, 5) = dt2.Recordset.Fields(19) ''��������

Excelapp.ActiveSheet.Cells(1, 8) = dt2.Recordset.Fields(16)  ''˾��ҵ��
Excelapp.ActiveSheet.Cells(2, 8) = dt2.Recordset.Fields(12) ''���ϵ�λ
Excelapp.ActiveSheet.Cells(11, 2) = dt2.Recordset.Fields(9) '''���λ��
End If

Excelapp.ActiveWindow.Zoom = 100   ' ���ô������ű���Ϊ 100%
     'Excelapp.Visible = True  ' ע�͵����� Excel Ӧ�ó���ɼ��Ĵ���
    Excelapp.DisplayAlerts = False   ' ������ʾ����

    Excelapp.ActiveSheet.PrintOut   ' ֱ�Ӵ�ӡ��ǰ������

    Set Excelapp = Nothing   ' �ͷ� Excel Ӧ�ó������
    Exit Sub   ' �˳��ӳ���

Ert:

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
