Attribute VB_Name = "称料系统"

 Type adwww6
  A As String * 30          ' 編碼         <30>          1    ~    30
  B As String * 6            ' 註 1         < 6 >        31    ~    36
  c As String * 6            ' 註 2          < 6 >       37   ~    42
  d As String * 6            ' 註 3          < 6 >   43   ~    48
  e As String * 30          ' 材質名稱   < 30 >  49   ~    78
  F As String * 10          ' 布重         < 10 >    79   ~   88
  g As String * 6            ' 浴比           < 6 >   89  ~    94
  h As String * 10          ' 總浴量       < 10 >   95   ~    104
  i As String * 12           ' 配方碼 ------  <  12  > 105  ~  116
  i2 As String * 6           '檔碼  200004   < 6 >   117  ~  122
  i3 As String * 1           '------  < 1 >   123   ~  123
  i4 As String * 1           '------  < 1 >   124   ~  124
  J(1 To 15) As String * 12  ' 代碼   < 12  *  15  = 180 >  125  ~  304
  k(1 To 15) As String * 8  ' 000.0000 配方   < 8  *  15  = 120 >  305  ~  424
  L(1 To 15) As String * 9  ' -----   00000.000 計量g   <9 * 15  = 135 >  425  ~  559
  L1(1 To 15) As String * 9  '-----   00000.000 實際計量g  < 9 * 15  = 135 >  560  ~  694
  M(1 To 15) As String * 1          '----   D/A   <  1  >  695  ~  695
  N(1 To 15) As String * 1          '----    %/g   <  1  >  696  ~   696
  o(1 To 15) As String * 4          '----濃度%   <  4  >   697  ~  700
  u As String * 8            ' 日時分 ----  <  8  >  701  ~  708
  W11 As String * 5          ' 檔碼 -----  <  5  >  709  ~   713
  W12 As String * 12           ' line \\  -----  <  12  >  714  ~  725
  DE As String * 1               '  ----  <  1  >   726  ~  726
  X As String * 2            ' 結束碼 chr(13)+chr(10)  <  2  >  727  ~   728
  End Type
  Public adwww6 As adwww6
  
Public Sub bpww666(S%, da$) '钡璸秖郎
   On Error GoTo ppccma
      
      namep$ = "\\ad1\c\adcc\DAT3\G" + da$ + ".666"
      
      op1% = FreeFile: Open namep$ For Random As #op1% Len = Len(adwww6)
      N& = LOF(op1%) / Len(adwww6) + 1
'     If S% = 0 Then s2& = N&
'      If S% = 1 Then Get #op1%, s2&, adwww6
      If S% = 2 Then Put #op1%, N&, adwww6
      
      Close #op1%
  Exit Sub
'================
ppccma:
    Close #op1%
End Sub
