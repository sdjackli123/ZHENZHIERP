Attribute VB_Name = "自动缩放"
Option Explicit
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private FormOldWidth As Long
'保存窗体的原始宽度
Private FormOldHeight As Long
'保存窗体的原始高度
Public fbl As Integer '''分辨率是否变化
Public suiping As Integer  '''当前水平分辨率
Public cuizhi As Integer  '''当前垂直分辨率
Public Declare Function SetParent Lib "user32.dll" ( _
          ByVal hWndChild As Long, _
          ByVal hWndNewParent As Long) As Long


'在调用ResizeForm前先调用本函数
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

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim pos(4) As Double
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    ScaleX = suiping / 1366
    ScaleY = cuizhi / 768
'    ScaleX = FormName.ScaleWidth / FormOldWidth
'    ScaleY = FormName.ScaleHeight / FormOldHeight
    '保存窗体宽度缩放比例
    '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
        '读取控件的原始位置与大小
        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
        pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
        Else
        pos(i) = 0
        End If
        '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
        Obj.Move pos(0) * ScaleX, pos(1) * ScaleY, pos(2) * ScaleX, pos(3) * ScaleY
    Next i
    Next Obj
    On Error GoTo 0
    
End Sub

Public Sub SetDeviceIndependentWindow(ThisForm As Form)
Dim Obj As Control   'Control 是一个对象, 表示所有Visual Basic 内部控件的类名
Dim DesignX As Integer  '代表设计系统的水平分辨率
Dim DesignY As Integer  ''代表设计系统的直分辨率

Dim XFactor As Single   '. 水平比例因子
Dim YFactor As Single   '. 垂直比例因子
Dim X As Integer
DesignX% = 1366: DesignY% = 768  '' . 假设设计时的分辨率为1366 * 768

 '计算当前屏幕尺寸与设计时使用的屏幕尺寸的比值
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
    '保存窗体宽度缩放比例
    ScaleY = FormName.ScaleHeight / FormOldHeight
    '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
        '读取控件的原始位置与大小
        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
        pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
        Else
        pos(i) = 0
        End If
        '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
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






