Attribute VB_Name = "��������"

Public Sub dmd(DT1 As Adodc, dt2 As Adodc, DH As String, DD As String)   ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�ȫ��,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����  FROM v_bmd WHERE ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  order by ƥ��"
DT1.Refresh
If DT1.Recordset.EOF Then
Excelapp.Quit
Set Excelapp = Nothing
End If

DT1.Recordset.MoveFirst
        Excelapp.ActiveSheet.Cells(4, 7) = Trim(DT1.Recordset.Fields(10)) '����
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(0) ''�ͻ�ȫ��
        Excelapp.ActiveSheet.Cells(6, 6) = DT1.Recordset.Fields(3) ''Ʒ��
        Excelapp.ActiveSheet.Cells(8, 2) = DT1.Recordset.Fields(2) ''���
        Excelapp.ActiveSheet.Cells(8, 4) = DT1.Recordset.Fields(1) ''����
        Excelapp.ActiveSheet.Cells(8, 6) = DT1.Recordset.Fields(8) ''���
        Excelapp.ActiveSheet.Cells(8, 8) = DT1.Recordset.Fields(5) ''ɫ��
        Excelapp.ActiveSheet.Cells(10, 2) = DT1.Recordset.Fields(4) ''��������
        Excelapp.ActiveSheet.Cells(10, 4) = DT1.Recordset.Fields(7) ''����
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

'DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from v_bmd where ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  group by ����"
'DT1.Refresh
dt2.RecordSource = "select * from ckgl where ���ݺ�='" & DH & "' "     ''���ﵥ�ݺű������DH���ܵ�������"
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
Excelapp.ActiveSheet.Cells(14, 1) = dt2.Recordset.Fields(15) ''������ϸ
Excelapp.ActiveSheet.Cells(32, 3) = dt2.Recordset.Fields(3) ''ë������
Excelapp.ActiveSheet.Cells(10, 6) = dt2.Recordset.Fields(4) ''ë��ƥ��
Excelapp.ActiveSheet.Cells(32, 6) = dt2.Recordset.Fields(3) ''ë������
Excelapp.ActiveSheet.Cells(10, 8) = dt2.Recordset.Fields(12) ''���ϵ�λ
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
Public Sub dmdms(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\mdms.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����,����  FROM bmd WHERE ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from bmd where ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub
Public Sub dmd100(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�ȫ��,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����  FROM v_bmd WHERE ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from v_bmd where ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub
Public Sub dmd100ms(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100ms.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����,����  FROM bmd WHERE ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from bmd where ����='" & DH & "' and ����='����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Public Sub xmdms(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\mdms.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����,����  FROM bmd WHERE ����='" & DH & "' and ����='С����' and �׺�='" & DD & "'  order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from bmd where ����='" & DH & "' and ����='С����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub
Public Sub xmd(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�ȫ��,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����,����  FROM v_bmd WHERE ����='" & DH & "' and ����='С����' and �׺�='" & DD & "' order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from v_bmd where ����='" & DH & "' and ����='С����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit


End Sub


Public Sub xmd100ms(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100ms.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����,����,����  FROM bmd WHERE ����='" & DH & "' and ����='С����' and �׺�='" & DD & "' order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from bmd where ����='" & DH & "' and ����='С����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub
Public Sub xmd100(DT1 As Adodc, DH As String, DD As String)    ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�ȫ��,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����  FROM v_bmd WHERE ����='" & DH & "' and ����='С����' and �׺�='" & DD & "' order by ƥ��"
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

DT1.RecordSource = "select ����,count(distinct ƥ��),sum(��������),sum(����) from v_bmd where ����='" & DH & "' and ����='С����' and �׺�='" & DD & "'  group by ����"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub


Public Sub dbq(DT1 As Adodc, DH As String, ph As Integer, DD As String, fs As String)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\bq.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT *  FROM v_bmd WHERE ����='" & DH & "' and ƥ��='" & ph & "'  and �׺�='" & DD & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        If IsNull(DT1.Recordset.Fields(18)) Then
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(0) '''�ͻ����
        Else
        Excelapp.ActiveSheet.Cells(3, 2) = DT1.Recordset.Fields(18) ''�ͻ�ȫ��
        End If
        Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(12) ''''ƥ��
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(1)  ''''����
        ''Excelapp.ActiveSheet.Cells(5, 5) = Format(DT1.Recordset.Fields(9), "#0.0") ''''����
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(8)    '''''ɫ��
        Excelapp.ActiveSheet.Cells(6, 5) = Format(DT1.Recordset.Fields(16), "#0.0")   '''''''����
        ''Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(2)    '''''���
        Excelapp.ActiveSheet.Cells(3, 5) = DT1.Recordset.Fields(5)    '''''����
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(3)    '''''Ʒ��
        Excelapp.ActiveSheet.Cells(4, 5) = DT1.Recordset.Fields(10)    '''''����
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(27)    '''''���
        Excelapp.ActiveSheet.Cells(5, 5) = Trim(DT1.Recordset.Fields(13))    '''''ʱ��
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(33) '''Ա����������
Excelapp.ActiveWindow.Zoom = 100

        Excelapp.Visible = True
        Excelapp.DisplayAlerts = False
        Excelapp.Sheets.PrintOut Copies:=fs   '''''��ӡ����
        Excelapp.Quit
        Set Excelapp = Nothing
        Exit Sub

Ert:

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing


End Sub
Public Sub dbqxs(DT1 As Adodc, DH As String, ph As Integer, DD As String, xh As Integer)   ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\bqxs.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT *  FROM v_bmd WHERE ����='" & DH & "' and ƥ��='" & ph & "' and ����='����' and ����='" & DD & "' and ���='" & xh & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        Excelapp.ActiveSheet.Cells(1, 5) = DT1.Recordset.Fields(19)
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(21)
        Excelapp.ActiveSheet.Cells(4, 5) = Trim(DT1.Recordset.Fields(6)) + "/" + Trim(DT1.Recordset.Fields(12))
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)
        Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(9)
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(8)
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(16)    '''''����
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

'Excelapp.Quit '�ر�EXCEL
Excelapp.Quit
Set Excelapp = Nothing


End Sub


Public Sub xbq(DT1 As Adodc, DH As String, DD As String, fs As Integer, fk As String)
    Dim i As Integer
    On Error GoTo Ert

    Dim Excelapp As Excel.Application
    Set Excelapp = New Excel.Application
    Excelapp.Visible = True ' ��Excel�ɼ�
    Excelapp.DisplayAlerts = False ' �ر�Excel������Ϣ

    ' ��ģ�幤����
    Dim wb As Excel.Workbook
    Set wb = Excelapp.Workbooks.Open(App.Path & "\��ӡģ��\����\bq.xls")
    Dim ws As Excel.Worksheet
    Set ws = wb.Sheets(1)

    DT1.RecordSource = "SELECT * FROM v_bmd WHERE ����='" & DH & "' AND �׺�='" & DD & "'"
    DT1.Refresh
    DT1.Recordset.MoveFirst

    ' ����ƥ�Ų������ӡ
    For i = 1 To fs
        With ws
            If IsNull(DT1.Recordset.Fields(18)) Then
                .Cells(3, 2) = DT1.Recordset.Fields(0)
            Else
                .Cells(3, 2) = DT1.Recordset.Fields(18)
            End If

            .Cells(7, 5) = i  '''' ƥ�ţ���1��fs
            .Cells(4, 2) = DT1.Recordset.Fields(1) '''' ����
            .Cells(5, 2) = DT1.Recordset.Fields(8) ''''' ɫ��
            .Cells(3, 5) = fk       ''''' ����
            .Cells(6, 2) = DT1.Recordset.Fields(3) ''''' Ʒ��
            .Cells(4, 5) = DT1.Recordset.Fields(10) ''''' ����
            .Cells(7, 2) = DT1.Recordset.Fields(27) ''''' ���
            .Cells(5, 5) = Trim(DT1.Recordset.Fields(13)) ''''' ʱ��
        End With
        ws.PrintOut Copies:=1, Collate:=True ' ��ӡ��ǰ������1��
    Next i
    
    ' ������˳�
    Excelapp.Quit
    Set Excelapp = Nothing
    Set wb = Nothing
    Exit Sub

Ert:
    ' ������ȷ��ExcelӦ����ȷ�ر�
    If Not Excelapp Is Nothing Then
        Excelapp.Quit
        Set Excelapp = Nothing
    End If
    If Not wb Is Nothing Then
        Set wb = Nothing
    End If
End Sub





Public Sub CLBB(Flex As VSFlexGrid, fd1, BT As String)  ''''�ޱ���

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
Excelapp.ActiveSheet.Cells(i, 1) = "�ϼ�"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub CLDY(Flex As VSFlexGrid, BT As String, Flex1 As VSFlexGrid)  ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\RSDY.xls")
'5)���õ�2��������Ϊ�������
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit
End Sub

Public Sub gzdc(Flex As VSFlexGrid, FD, BT As String)   ''''�ޱ���

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

Excelapp.ActiveSheet.Cells(i, 1) = "�ϼƽ��"
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



Public Sub dmd100dc(DT1 As Adodc, DH As String, pm As String)   ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����  FROM bmd WHERE ����='" & DH & "' and ����='����' and Ʒ��='" & pm & "' order by ƥ��"
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

DT1.RecordSource = "select distinct Ʒ��,���߷���,����,ƥ�� from bmd where ����='" & DH & "' and ����='����' and Ʒ��='" & pm & "'"
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

DT1.RecordSource = "select distinct ��ע from kpd where ����='" & DH & "' and Ʒ��='" & pm & "'"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Public Sub xmd100dc(DT1 As Adodc, DH As String, pm As String)   ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\md100.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,���,Ʒ��,���߷���,ɫ��,��������,����,���,ƥ��,����  FROM bmd WHERE ����='" & DH & "' and ����='С����' and Ʒ��='" & pm & "' order by ƥ��"
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

DT1.RecordSource = "select distinct Ʒ��,���߷���,����,ƥ�� from bmd where ����='" & DH & "' and ����='С����' and Ʒ��='" & pm & "'"
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

DT1.RecordSource = "select distinct ��ע from kpd where ����='" & DH & "' and Ʒ��='" & pm & "'"
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

'Excelapp.Quit '�ر�EXCEL
Set Excelapp = Nothing
Excelapp.Quit

End Sub

Public Sub mpbq(DT1 As Adodc, DH As String, xh As Integer)     ''''�ޱ���

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
Excelapp.Workbooks.Open (App.Path & "\��ӡģ��\����\mpbq.xls")
'5)���õ�2��������Ϊ�������
        Excelapp.Sheets(1).Activate

DT1.RecordSource = "SELECT �ͻ�����,����,ɫ��,��ǩ,���߷���,Ʒ��,����Ҫ��  FROM kpd WHERE ����='" & DH & "' and ip='" & xh & "'"
DT1.Refresh

i = 0
'For i = 0 To MN - 1
        Excelapp.ActiveSheet.Cells(4, 2) = DT1.Recordset.Fields(0)  '''�ͻ�
'        Excelapp.ActiveSheet.Cells(4, 5) = dt1.Recordset.Fields(2)
        Excelapp.ActiveSheet.Cells(5, 2) = DT1.Recordset.Fields(1)  ''''�׺�
        Excelapp.ActiveSheet.Cells(5, 5) = DT1.Recordset.Fields(2)  ''''�׺�
        Excelapp.ActiveSheet.Cells(6, 2) = DT1.Recordset.Fields(3)  ''''���
        Excelapp.ActiveSheet.Cells(6, 5) = DT1.Recordset.Fields(4)  ''''����
        Excelapp.ActiveSheet.Cells(7, 2) = DT1.Recordset.Fields(5)  ''''Ʒ��
        Excelapp.ActiveSheet.Cells(7, 5) = DT1.Recordset.Fields(6)  ''''����
        

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
Public Sub mbdy(DT1 As Adodc, selectedGuoHao As String, sh As Excel.Worksheet)
    On Error GoTo errorhandler

    ' ȷ�� DT1 ����ȷ���ӵ����ݿ�
    If DT1.ConnectionString = "" Then
        MsgBox "DT1 is not connected to the database."
        Exit Sub
    End If

    ' ��ѯ���ݿ��ȡ�ͻ������ڡ�ɫ���ɫ��
    Debug.Print "Executing SQL query through DT1" ' ��ӡ������Ϣ
    DT1.RecordSource = "SELECT �ͻ�����, ����, ɫ��, ɫ�� FROM v_kpd_khmb WHERE ����='" & selectedGuoHao & "'"
    DT1.Refresh

    ' ȷ����¼������
    If DT1.Recordset Is Nothing Then
        MsgBox "SQL query failed. Check the database connection and query."
        GoTo Cleanup
    End If

    ' ��ʼ�� rowIndex
    Dim rowIndex As Integer
    rowIndex = sh.Cells(sh.Rows.count, 1).End(xlUp).Row + 1
    Debug.Print "Row index initialized to: " & rowIndex ' ��ӡ������Ϣ

    If Not DT1.Recordset.EOF Then
        Debug.Print "Record found for GuoHao: " & selectedGuoHao ' ��ӡ������Ϣ
        sh.Cells(rowIndex, 1).value = DT1.Recordset.Fields("�ͻ�����").value
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "���ڣ�" & DT1.Recordset.Fields("����").value & " ������ϸ"
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "��ɫ��" & DT1.Recordset.Fields("ɫ��").value & " ɫ�ţ�" & DT1.Recordset.Fields("ɫ��").value
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "������Ϣ�������뼰ʱ��֪����ȷ��ɫ��"
        rowIndex = rowIndex + 1
        sh.Cells(rowIndex, 1).value = "" ' ��һ�У�ȷ����Ϣ֮���п���
        rowIndex = rowIndex + 1
    Else
        Debug.Print "No records found for GuoHao: " & selectedGuoHao ' ��ӡ������Ϣ
    End If

Cleanup:
    Exit Sub

errorhandler:
    MsgBox "����: " & Err.Description
    Resume Cleanup
End Sub

