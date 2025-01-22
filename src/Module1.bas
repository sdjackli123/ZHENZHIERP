Attribute VB_Name = "Module1"
Public Password As String
Public dat1, dat2 As Date:
Public dat3 As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public ypxx As String
Public Sub entertotab(keyasc As Integer)
If keyasc = 13 Then
SendKeys "{tab}"
End If
End Sub

