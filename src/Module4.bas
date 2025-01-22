Attribute VB_Name = "Module4"
Public Sub CPFH(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data9.RecordSource = "SELECT * FROM CLYSHZ order BY 加工单位"
Data9.Refresh
Data8.RecordSource = "SELECT * FROM CLZZPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLZZPZ WHERE 日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data7.Refresh
PZH = "5-1"
If Data7.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data7.Recordset.Fields(0) + 1)
End If

If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "主营收入"
Data8.Recordset.Fields(1) = "应收账款"
Data8.Recordset.Fields(2) = Data9.Recordset.Fields(1)
Data8.Recordset.Fields(3) = "主营业务收入"
Data8.Recordset.Fields(4) = ""
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(4)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = dt3
Data8.Recordset.Fields(8) = ""
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = "自动"
Data8.Recordset.Update

Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("成品发货单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.凭证号,3))) FROM CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "5-1"
If Data7.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("成品发货单转账成功！" + "生成" + Str(KLLLL) + "凭证")

End If
End Sub

Public Sub FKHZXJ()   '''''''''付款汇总--现金
On Error Resume Next

Data8.RecordSource = "SELECT * FROM CLFKPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'2-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "2-1"
If Not Data7.Recordset.EOF Then
PZH = "2-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM JGZCX1 where val(本期现付款)>0"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "购材料"     '''''摘要
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)      ''''借方总账科目
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)         ''''借方明细科目
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(4)
End If
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)     '''''''''''贷方总账科目
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)       '''贷方明细科目
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(0)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(6)     ''''''''发生金额
Data8.Recordset.Fields(6) = PZH                           '''''''''''''凭证号
Data8.Recordset.Fields(7) = DTPicker3.Value               '''''''操作日期
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""                           ''''''''''记账
Data8.Recordset.Fields(10) = ""                          ''''''''''''复核
Data8.Recordset.Fields(11) = DBCombo3.Text       ''''''''''''''制单
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "已"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "现金付款单"
Data2.Recordset.Fields(3) = "付款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'2-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "2-1"
If Not Data7.Recordset.EOF Then
PZH = "2-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金付款凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "现金付款单"
Data2.Recordset.Fields(3) = "付款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
End If
End Sub
''''''''''''''
Public Sub FKHZYH()   '''''''''付款汇总---银行存款
On Error Resume Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'4-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "4-1"
If Not Data7.Recordset.EOF Then
PZH = "4-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (审核确认<>'已' OR 审核确认=NULL) AND 贷方金额<>'0' AND 借方金额='0' AND INSTR(类别,'银行存款')>0 AND (审核<>'已' OR 审核=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(4)
End If
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(0)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(6)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "已"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款付款凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "银行付款单"
Data2.Recordset.Fields(3) = "付款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLFKPZ WHERE INSTR(凭证号,'4-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "4-1"
If Not Data7.Recordset.EOF Then
PZH = "4-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "银行付款单"
Data2.Recordset.Fields(3) = "付款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款付款凭证")
End If
End Sub

Public Sub SKHZXJ()    ''''''''收款汇总----现金
On Error Resume Next
If DBCombo3.Text = "" Then
MsgBox ("请选择复核员")
Exit Sub
End If
Data8.RecordSource = "SELECT * FROM CLSKPZ"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'1-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "1-1"
If Not Data7.Recordset.EOF Then
PZH = "1-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (审核确认<>'已' OR 审核确认=NULL) AND 借方金额<>'0' AND 贷方金额='0' AND INSTR(类别,'现金')>0 AND (审核<>'已' OR 审核=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(0)
End If
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(4)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(5)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "已"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "现金收款单"
Data2.Recordset.Fields(3) = "收款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'1-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "1-1"
If Not Data7.Recordset.EOF Then
PZH = "1-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "现金收款单"
Data2.Recordset.Fields(3) = "收款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "现金收款凭证")
End If
End Sub

Public Sub SKHZYH()    ''''''''收款汇总----银行存款
On Error Resume Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'3-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "3-1"
If Not Data7.Recordset.EOF Then
PZH = "3-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Data9.RecordSource = "SELECT * FROM TZJZMX WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "') AND (审核确认<>'已' OR 审核确认=NULL) AND 借方金额<>'0' AND 贷方金额='0' AND INSTR(类别,'银行存款')>0 AND (审核<>'已' OR 审核=NULL)"
Data9.Refresh
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = Data9.Recordset.Fields(3)
If InStr(Data9.Recordset.Fields(0), "-") > 0 Then
Data8.Recordset.Fields(1) = Left(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") - 1)
Data8.Recordset.Fields(2) = Mid(Data9.Recordset.Fields(0), InStr(Data9.Recordset.Fields(0), "-") + 1)
Else
Data8.Recordset.Fields(1) = Data9.Recordset.Fields(0)
End If
If InStr(Data9.Recordset.Fields(4), "-") > 0 Then
Data8.Recordset.Fields(3) = Left(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") - 1)
Data8.Recordset.Fields(4) = Mid(Data9.Recordset.Fields(4), InStr(Data9.Recordset.Fields(4), "-") + 1)
Else
Data8.Recordset.Fields(3) = Data9.Recordset.Fields(4)
End If
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(5)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = Data9.Recordset.Fields(2)
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.Edit    '''''''''''''''''''''''''''''''
Data9.Recordset.Fields(12) = "已"
Data9.Recordset.Update    ''''''''''''''''''''''''''''''
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款收款凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "银行收款单"
Data2.Recordset.Fields(3) = "收款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
Data7.RecordSource = "SELECT MAX(VAL(MID(凭证号,3))) FROM CLSKPZ WHERE INSTR(凭证号,'3-')>0 AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "3-1"
If Not Data7.Recordset.EOF Then
PZH = "3-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
KLLLL = KLLLL + 1
Loop
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "银行收款单"
Data2.Recordset.Fields(3) = "收款凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "银行存款收款凭证")
End If
End Sub



Public Sub CLRK(dt1 As Date, dt2 As Date, dt3 As Date)
On Error Resume Next

Data9.RecordSource = "SELECT * FROM JGZCX1 where val(本期应付款)>0"
Data9.Refresh
If Data9.Recordset.EOF Then Exit Sub
Data10.RecordSource = "SELECT * FROM CLZZPZ"
Data10.Refresh
Data11.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.凭证号,3))) FROM CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "5-1"
If Data11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data10.Recordset.AddNew
Data10.Recordset.Fields(0) = "购材料"
Data10.Recordset.Fields(1) = "库存物资"
Data10.Recordset.Fields(2) = ""
Data10.Recordset.Fields(3) = "应付账款"
Data10.Recordset.Fields(4) = Data9.Recordset.Fields(0)
Data10.Recordset.Fields(5) = Data9.Recordset.Fields(2)
Data10.Recordset.Fields(6) = PZH
Data10.Recordset.Fields(7) = CDate(dt3)
Data10.Recordset.Fields(8) = ""
Data10.Recordset.Fields(9) = ""
Data10.Recordset.Fields(10) = ""
Data10.Recordset.Fields(11) = "自动-材料"
Data10.Recordset.Update
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data11.RecordSource = "SELECT MAX(VAL(MID(CLZZPZ.凭证号,3))) FROM CLZZPZ WHERE CLZZPZ.日期 BETWEEN CDATE('" & dt1 & "') AND CDATE('" & dt2 & "')"
Data11.Refresh
PZH = "5-1"
If Data11.Recordset.EOF Then
PZH = "5-1"
Else
PZH = "5-" + Trim(Data11.Recordset.Fields(0) + 1)
End If
Loop
MsgBox ("材料入库单转账成功！" + "生成" + Str(KLLLL) + "凭证")
End Sub

Public Sub CLCK()
On Error Resume Next
Data8.Database.Execute "DELETE * FROM CLSCHZ"
lo = "d:\数据库\ssdt\" + LJB + "\zcw.MDB"       '''''''''''''''''''''''经典
Data3.Database.Execute "INSERT INTO CLSCHZ(金额) IN'" & lo & "' SELECT FORMAT(SUM(合计金额),'#0.00') AS 金额 FROM KPD WHERE (审核<>'已' OR 审核=NULL) AND 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data8.Database.Execute "DELETE * FROM CLSCHZ WHERE 金额=NULL"
Data9.RecordSource = "SELECT * FROM CLSCHZ"
Data9.Refresh
Data8.RecordSource = "SELECT * FROM CLSCCB"
Data8.Refresh
Data7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.凭证号,3))) FROM CLSCCB WHERE CLSCCB.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "S-1"
If Not Data7.Recordset.EOF Then
PZH = "S-" + Trim(Data7.Recordset.Fields(0) + 1)
Else
PZH = "S-1"
End If
If Data9.Recordset.EOF Then
Exit Sub
Else
Data9.Recordset.MoveFirst
KLLLL = 1
Do While Not Data9.Recordset.EOF
For i = 1 To 5
Data8.Recordset.AddNew
Data8.Recordset.Fields(0) = "耗用原材料"
Data8.Recordset.Fields(1) = "生产成本"
Data8.Recordset.Fields(2) = "直接生产成本"
Data8.Recordset.Fields(3) = "库存物资"
Data8.Recordset.Fields(4) = ""
Data8.Recordset.Fields(5) = Data9.Recordset.Fields(4)
Data8.Recordset.Fields(6) = PZH
Data8.Recordset.Fields(7) = DTPicker3.Value
Data8.Recordset.Fields(8) = ""
Data8.Recordset.Fields(9) = ""
Data8.Recordset.Fields(10) = ""
Data8.Recordset.Fields(11) = DBCombo3.Text
Data8.Recordset.Update
Data9.Recordset.MoveNext
If Data9.Recordset.EOF Then
Data4.Database.Execute "UPDATE KPD SET 审核='已' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "成本凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "材料出库单"
Data2.Recordset.Fields(3) = "成本凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
Exit Sub
End If
Next
KLLLL = KLLLL + 1
Data7.RecordSource = "SELECT MAX(VAL(MID(CLSCCB.凭证号,3))) FROM CLSCCB WHERE CLSCCB.日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
Data7.Refresh
PZH = "S-1"
If Data7.Recordset.EOF Then
PZH = "S-1"
Else
PZH = "S-" + Trim(Data7.Recordset.Fields(0) + 1)
End If
Loop
Data4.Database.Execute "UPDATE KPD SET 审核='已' WHERE 日期 BETWEEN CDATE('" & Text1.Text & "') AND CDATE('" & Text2.Text & "')"
MsgBox ("转账成功！" + "生成" + Str(KLLLL) + "成本凭证")
Data2.Recordset.AddNew
Data2.Recordset.Fields(0) = DTPicker3.Value
Data2.Recordset.Fields(1) = Text1.Text
Data2.Recordset.Fields(2) = "材料出库单"
Data2.Recordset.Fields(3) = "成本凭证"
Data2.Recordset.Fields(4) = Str(KLLLL)
Data2.Recordset.Update
Data2.Refresh
End If
End Sub

