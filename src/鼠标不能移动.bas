Attribute VB_Name = "��겻���ƶ�"
Option Explicit
Dim r As RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)    '''���̶�ȡ��

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long  '''���̶�

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long           '''''�������
Dim blnShow As Boolean
