Attribute VB_Name = "鼠标不能移动"
Option Explicit
Dim r As RECT
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)    '''鼠标固定取消

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long  '''鼠标固定

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long           '''''鼠标隐藏
Dim blnShow As Boolean
