Attribute VB_Name = "�Զ�����"
Option Explicit
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private FormOldWidth As Long
'���洰���ԭʼ���
Private FormOldHeight As Long
'���洰���ԭʼ�߶�
Public fbl As Integer '''�ֱ����Ƿ�仯
Public suiping As Integer  '''��ǰˮƽ�ֱ���
Public cuizhi As Integer  '''��ǰ��ֱ�ֱ���
Public Declare Function SetParent Lib "user32.dll" ( _
          ByVal hWndChild As Long, _
          ByVal hWndNewParent As Long) As Long


'�ڵ���ResizeFormǰ�ȵ��ñ�����
'Public Sub ResizeInit()Sub ResizeInit(FormName As Form)
Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
    Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub

'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim pos(4) As Double
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    ScaleX = suiping / 1366
    ScaleY = cuizhi / 768
'    ScaleX = FormName.ScaleWidth / FormOldWidth
'    ScaleY = FormName.ScaleHeight / FormOldHeight
    '���洰�������ű���
    '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
        '��ȡ�ؼ���ԭʼλ�����С
        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
        pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
        Else
        pos(i) = 0
        End If
        '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
        Obj.Move pos(0) * ScaleX, pos(1) * ScaleY, pos(2) * ScaleX, pos(3) * ScaleY
    Next i
    Next Obj
    On Error GoTo 0
    
End Sub

Public Sub SetDeviceIndependentWindow(ThisForm As Form)
Dim Obj As Control   'Control ��һ������, ��ʾ����Visual Basic �ڲ��ؼ�������
Dim DesignX As Integer  '�������ϵͳ��ˮƽ�ֱ���
Dim DesignY As Integer  ''�������ϵͳ��ֱ�ֱ���

Dim XFactor As Single   '. ˮƽ��������
Dim YFactor As Single   '. ��ֱ��������
Dim X As Integer
DesignX% = 1366: DesignY% = 768  '' . �������ʱ�ķֱ���Ϊ1366 * 768

 '���㵱ǰ��Ļ�ߴ������ʱʹ�õ���Ļ�ߴ�ı�ֵ
XFactor = (Screen.Width / Screen.TwipsPerPixelX) / DesignX
YFactor = (Screen.Height / Screen.TwipsPerPixelY) / DesignY

If XFactor = 1 And YFactor = 1 Then
fbl = 0
Else: fbl = 1
End If
End Sub

Public Sub zdResizeForm(FormName As Form)
    Dim pos(4) As Double
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
   
    ScaleX = FormName.ScaleWidth / FormOldWidth
    '���洰�������ű���
    ScaleY = FormName.ScaleHeight / FormOldHeight
    '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
        '��ȡ�ؼ���ԭʼλ�����С
        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
        pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
        Else
        pos(i) = 0
        End If
        '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
        Obj.Move pos(0) * ScaleX, pos(1) * ScaleY, pos(2) * ScaleX, pos(3) * ScaleY
    Next i
    Next Obj
    On Error GoTo 0
End Sub

Public Sub zdResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
    Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub






